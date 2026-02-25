#!/usr/bin/env node
/**
 * convert_mermaid_to_excalidraw.mjs
 *
 * Extracts ```mermaid code blocks from a Markdown file,
 * converts each to:
 *   1. A standalone .excalidraw (JSON) file   — editable in VS Code Excalidraw extension
 *   2. A Mermaid-rendered .svg file           — viewable everywhere
 * Then optionally replaces the mermaid code blocks in the Markdown with image references.
 *
 * Usage:
 *   cd scripts
 *   npm install
 *   npm run install:browsers          # one-time: download Chromium for Playwright
 *   node convert_mermaid_to_excalidraw.mjs [--replace] <markdown-file>
 *
 * Options:
 *   --replace   Replace mermaid code blocks in the markdown with ![](diagrams/...) image refs
 *               Without this flag, diagrams are generated but the markdown is left unchanged.
 *
 * Output:
 *   diagrams/
 *     01-<slug>.excalidraw        # Native Excalidraw JSON (editable)
 *     01-<slug>.svg               # Mermaid-rendered SVG   (viewable)
 */

import { chromium } from "playwright";
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";

// ─── Helpers ────────────────────────────────────────────────────────────────

/** Generate a short random id. */
function rid() {
  return crypto.randomBytes(8).toString("hex");
}

/** Extract all ```mermaid ... ``` blocks from markdown text. */
function extractMermaidBlocks(markdown) {
  const regex = /```mermaid\n([\s\S]*?)```/g;
  const blocks = [];
  let match;
  while ((match = regex.exec(markdown)) !== null) {
    blocks.push({
      fullMatch: match[0],
      code: match[1].trim(),
      index: match.index,
    });
  }
  return blocks;
}

/** Derive a short slug from the mermaid code. */
function slugify(code, index) {
  const subgraphMatch = code.match(/subgraph\s+\w+\["([^"]+)"/);
  const titleMatch = code.match(/---\s*title:\s*(.+)/);

  let name;
  if (subgraphMatch) {
    name = subgraphMatch[1];
  } else if (titleMatch) {
    name = titleMatch[1];
  } else {
    const typeMatch = code.match(/^(\w+)/);
    name = typeMatch ? typeMatch[1] : "diagram";
  }

  return name
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 40);
}

// ─── Skeleton → Native Excalidraw Conversion ────────────────────────────────
//
// The @excalidraw/mermaid-to-excalidraw library returns "skeleton" elements
// with non-standard properties:
//   - Shapes use `label: { text, fontSize, ... }` for bound text
//   - Arrows use `start: { id }` and `end: { id }` for bindings
//
// Excalidraw's native format requires:
//   - Separate `type: "text"` elements with `containerId` pointing to parent
//   - Parent shapes have `boundElements: [{ id: textId, type: "text" }]`
//   - Arrows use `startBinding: { elementId, focus, gap }` and `endBinding`
//   - Target shapes of arrows also list the arrow in their `boundElements`
//
// This function performs that conversion without needing the React-based
// `convertToExcalidrawElements` from @excalidraw/excalidraw.
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Default Excalidraw element properties.
 */
function excalidrawDefaults() {
  return {
    fillStyle: "solid",
    strokeWidth: 2,
    strokeStyle: "solid",
    roughness: 1,
    opacity: 100,
    angle: 0,
    strokeColor: "#1e1e1e",
    backgroundColor: "transparent",
    seed: Math.floor(Math.random() * 2_000_000_000),
    version: 1,
    versionNonce: Math.floor(Math.random() * 2_000_000_000),
    isDeleted: false,
    frameId: null,
    link: null,
    locked: false,
    updated: Date.now(),
  };
}

/**
 * Measure approximate text width (monospace heuristic: 0.6 * fontSize * maxLineLength).
 */
function measureText(text, fontSize) {
  const lines = text.split("\n");
  const maxLen = Math.max(...lines.map((l) => l.length));
  const width = Math.ceil(maxLen * fontSize * 0.6) + 20;
  const lineHeight = fontSize * 1.35;
  const height = Math.ceil(lines.length * lineHeight);
  return { width, height };
}

/**
 * Convert skeleton elements from mermaid-to-excalidraw into
 * native Excalidraw elements with proper bound text elements.
 */
function convertSkeletonToNative(skeletonElements) {
  const nativeElements = [];
  // Map from original element id → native element (for arrow binding fixup)
  const elById = new Map();

  for (const skel of skeletonElements) {
    const isShape =
      skel.type === "rectangle" ||
      skel.type === "ellipse" ||
      skel.type === "diamond";
    const isArrow = skel.type === "arrow";
    const isLine = skel.type === "line";
    const isText = skel.type === "text";

    if (isShape) {
      // ── Shape element ───────────────────────────────────────────────
      const shape = {
        ...excalidrawDefaults(),
        id: skel.id,
        type: skel.type,
        x: skel.x ?? 0,
        y: skel.y ?? 0,
        width: skel.width ?? 100,
        height: skel.height ?? 60,
        groupIds: skel.groupIds ?? [],
        roundness: skel.roundness ?? { type: 3 },
        boundElements: [],
        index: skel.index ?? undefined,
      };

      // Preserve any explicitly set properties from skeleton
      if (skel.strokeWidth != null) shape.strokeWidth = skel.strokeWidth;
      if (skel.strokeColor) shape.strokeColor = skel.strokeColor;
      if (skel.backgroundColor) shape.backgroundColor = skel.backgroundColor;
      if (skel.strokeStyle) shape.strokeStyle = skel.strokeStyle;
      if (skel.fillStyle) shape.fillStyle = skel.fillStyle;

      nativeElements.push(shape);
      elById.set(shape.id, shape);

      // Create bound text element if label exists
      if (skel.label && skel.label.text) {
        const rawText = skel.label.text
          .replace(/<br\s*\/?>/gi, "\n")
          .replace(/<[^>]+>/g, "");
        const fontSize = skel.label.fontSize ?? 16;
        const verticalAlign = skel.label.verticalAlign ?? "middle";
        const textAlign = skel.label.textAlign ?? "center";
        const { width: tw, height: th } = measureText(rawText, fontSize);

        const textId = `${skel.id}_label_${rid()}`;

        // Position text centered within the shape
        const textX = shape.x + (shape.width - tw) / 2;
        let textY;
        if (verticalAlign === "top") {
          textY = shape.y + 4;
        } else {
          textY = shape.y + (shape.height - th) / 2;
        }

        const textEl = {
          ...excalidrawDefaults(),
          id: textId,
          type: "text",
          x: textX,
          y: textY,
          width: tw,
          height: th,
          groupIds: skel.label.groupIds ?? skel.groupIds ?? [],
          roundness: null,
          boundElements: null,
          text: rawText,
          originalText: rawText,
          autoResize: true,
          fontSize,
          fontFamily: 1, // Virgil (hand-drawn)
          textAlign,
          verticalAlign,
          containerId: shape.id,
          lineHeight: 1.25,
          index: skel.index ? skel.index + "_t" : undefined,
        };

        // Link text to shape
        shape.boundElements.push({ id: textId, type: "text" });

        nativeElements.push(textEl);
      }
    } else if (isArrow || isLine) {
      // ── Arrow / Line element ────────────────────────────────────────
      const arrow = {
        ...excalidrawDefaults(),
        id: skel.id,
        type: skel.type,
        x: skel.x ?? 0,
        y: skel.y ?? 0,
        width: skel.width ?? 0,
        height: skel.height ?? 0,
        groupIds: skel.groupIds ?? [],
        roundness: skel.roundness ?? { type: 2 },
        boundElements: [],
        points: skel.points ?? [[0, 0], [100, 0]],
        lastCommittedPoint: null,
        startBinding: null,
        endBinding: null,
        startArrowhead: skel.startArrowhead ?? null,
        endArrowhead: isArrow ? (skel.endArrowhead ?? "arrow") : null,
        index: skel.index ?? undefined,
      };

      if (skel.strokeWidth != null) arrow.strokeWidth = skel.strokeWidth;
      if (skel.strokeColor) arrow.strokeColor = skel.strokeColor;
      if (skel.strokeStyle) arrow.strokeStyle = skel.strokeStyle;

      // Skeleton uses `start: { id }` and `end: { id }` for bindings
      if (skel.start?.id) {
        arrow.startBinding = {
          elementId: skel.start.id,
          focus: 0,
          gap: 5,
          fixedPoint: null,
        };
      }
      if (skel.end?.id) {
        arrow.endBinding = {
          elementId: skel.end.id,
          focus: 0,
          gap: 5,
          fixedPoint: null,
        };
      }

      nativeElements.push(arrow);
      elById.set(arrow.id, arrow);

      // Create bound text element for arrow label
      if (skel.label && skel.label.text) {
        const rawText = skel.label.text
          .replace(/<br\s*\/?>/gi, "\n")
          .replace(/<[^>]+>/g, "");
        const fontSize = skel.label.fontSize ?? 16;
        const { width: tw, height: th } = measureText(rawText, fontSize);

        const textId = `${skel.id}_label_${rid()}`;

        // Position text near the midpoint of the arrow
        const pts = arrow.points;
        const midIdx = Math.floor(pts.length / 2);
        const midPt = pts[midIdx] ?? [0, 0];
        const textX = arrow.x + midPt[0] - tw / 2;
        const textY = arrow.y + midPt[1] - th / 2;

        const textEl = {
          ...excalidrawDefaults(),
          id: textId,
          type: "text",
          x: textX,
          y: textY,
          width: tw,
          height: th,
          groupIds: skel.label.groupIds ?? [],
          roundness: null,
          boundElements: null,
          text: rawText,
          originalText: rawText,
          autoResize: true,
          fontSize,
          fontFamily: 1,
          textAlign: "center",
          verticalAlign: "middle",
          containerId: arrow.id,
          lineHeight: 1.25,
          index: skel.index ? skel.index + "_t" : undefined,
        };

        arrow.boundElements.push({ id: textId, type: "text" });
        nativeElements.push(textEl);
      }
    } else if (isText) {
      // ── Standalone text element ─────────────────────────────────────
      const rawText = (skel.text ?? "")
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<[^>]+>/g, "");
      const fontSize = skel.fontSize ?? 16;
      const { width: tw, height: th } = measureText(rawText, fontSize);

      nativeElements.push({
        ...excalidrawDefaults(),
        id: skel.id,
        type: "text",
        x: skel.x ?? 0,
        y: skel.y ?? 0,
        width: skel.width ?? tw,
        height: skel.height ?? th,
        groupIds: skel.groupIds ?? [],
        roundness: null,
        boundElements: null,
        text: rawText,
        originalText: rawText,
        autoResize: true,
        fontSize,
        fontFamily: 1,
        textAlign: skel.textAlign ?? "center",
        verticalAlign: "middle",
        containerId: null,
        lineHeight: 1.25,
        index: skel.index ?? undefined,
      });
    } else {
      // Unknown type — pass through with defaults
      nativeElements.push({
        ...excalidrawDefaults(),
        ...skel,
        boundElements: skel.boundElements ?? [],
      });
    }
  }

  // ── Fix-up pass: add arrow references to target shapes' boundElements ──
  for (const el of nativeElements) {
    if (el.type !== "arrow" && el.type !== "line") continue;

    if (el.startBinding?.elementId) {
      const target = elById.get(el.startBinding.elementId);
      if (target?.boundElements) {
        const alreadyLinked = target.boundElements.some(
          (b) => b.id === el.id && b.type === "arrow"
        );
        if (!alreadyLinked) {
          target.boundElements.push({ id: el.id, type: "arrow" });
        }
      }
    }
    if (el.endBinding?.elementId) {
      const target = elById.get(el.endBinding.elementId);
      if (target?.boundElements) {
        const alreadyLinked = target.boundElements.some(
          (b) => b.id === el.id && b.type === "arrow"
        );
        if (!alreadyLinked) {
          target.boundElements.push({ id: el.id, type: "arrow" });
        }
      }
    }
  }

  // Remove undefined fields
  for (const el of nativeElements) {
    for (const key of Object.keys(el)) {
      if (el[key] === undefined) delete el[key];
    }
  }

  return nativeElements;
}

/**
 * Build the self-contained HTML page that loads Mermaid + mermaid-to-excalidraw
 * from CDN (esm.sh) and exposes global conversion functions.
 */
function buildConverterHTML() {
  return `<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body>
<div id="mermaid-container"></div>
<script type="module">
  import mermaid from "https://esm.sh/mermaid@11?bundle";
  import { parseMermaidToExcalidraw } from "https://esm.sh/@excalidraw/mermaid-to-excalidraw@0.3?bundle";

  mermaid.initialize({
    startOnLoad: false,
    theme: "default",
    securityLevel: "loose",
    fontFamily: "sans-serif",
  });

  // ── Render Mermaid code to SVG string ──────────────────────────────
  window.renderMermaidSvg = async (code, id) => {
    try {
      const { svg } = await mermaid.render(id, code);
      const container = document.getElementById("mermaid-container");
      container.innerHTML = svg;
      return { ok: true, svg };
    } catch (e) {
      return { ok: false, error: e.message };
    }
  };

  // ── Convert Mermaid code to Excalidraw skeleton elements ───────────
  window.convertMermaidToExcalidraw = async (code) => {
    try {
      const { elements, files } = await parseMermaidToExcalidraw(code, {
        fontSize: 16,
      });
      return { ok: true, elements, files: files || {} };
    } catch (e) {
      return { ok: false, error: e.message };
    }
  };

  window.__ready = true;
</script>
</body>
</html>`;
}

/**
 * Build a full .excalidraw scene JSON from native elements.
 */
function buildExcalidrawScene(nativeElements, files = {}) {
  return {
    type: "excalidraw",
    version: 2,
    source: "mermaid-to-excalidraw",
    elements: nativeElements,
    appState: {
      gridSize: null,
      viewBackgroundColor: "#ffffff",
    },
    files,
  };
}

/**
 * Post-process SVG in the browser: replace all <foreignObject> elements
 * (which only render in browsers) with native SVG <text> elements so the
 * exported SVG displays correctly everywhere (VS Code, GitHub, image viewers).
 */
async function postProcessSvgInBrowser(page) {
  return await page.evaluate(() => {
    const svgEl = document.querySelector("#mermaid-container svg");
    if (!svgEl) return { ok: false, error: "No SVG in container" };

    const NS = "http://www.w3.org/2000/svg";
    const foreignObjects = svgEl.querySelectorAll("foreignObject");

    for (const fo of foreignObjects) {
      const spans = fo.querySelectorAll(
        "span.nodeLabel, span.edgeLabel, span.cluster-label, span"
      );
      let textContent = "";
      if (spans.length > 0) {
        const best = spans[spans.length - 1];
        textContent = best.textContent.trim();
      } else {
        textContent = fo.textContent.trim();
      }

      if (!textContent) {
        fo.parentNode.removeChild(fo);
        continue;
      }

      const htmlContent = fo.innerHTML;
      let lines;
      const pTags = fo.querySelectorAll("p");
      if (pTags.length > 1) {
        lines = Array.from(pTags)
          .map((p) => p.textContent.trim())
          .filter(Boolean);
      } else if (/<br\s*\/?>/i.test(htmlContent)) {
        lines = textContent.split(/\n/).filter(Boolean);
        if (lines.length === 1) {
          const div = document.createElement("div");
          div.innerHTML = htmlContent;
          const raw = div.innerHTML.replace(/<br\s*\/?>/gi, "\n");
          lines = raw
            .replace(/<[^>]+>/g, "")
            .split("\n")
            .map((s) => s.trim())
            .filter(Boolean);
        }
      } else {
        lines = [textContent];
      }

      const foWidth = parseFloat(fo.getAttribute("width")) || 0;
      const foHeight = parseFloat(fo.getAttribute("height")) || 0;

      let fontSize = 14;
      let fontWeight = "normal";
      const fill = "#333";
      const isClusterLabel = fo.closest(".cluster-label") !== null;

      if (isClusterLabel) fontWeight = "bold";

      const textEl = document.createElementNS(NS, "text");
      textEl.setAttribute("dominant-baseline", "central");
      textEl.setAttribute("text-anchor", "middle");
      textEl.setAttribute("fill", fill);
      textEl.setAttribute("font-family", "sans-serif");
      textEl.setAttribute("font-size", fontSize);
      if (fontWeight !== "normal")
        textEl.setAttribute("font-weight", fontWeight);

      const lineHeight = fontSize * 1.4;
      const totalTextHeight = lines.length * lineHeight;
      const startY = (foHeight - totalTextHeight) / 2 + lineHeight / 2;
      const cx = foWidth / 2;

      for (let li = 0; li < lines.length; li++) {
        const tspan = document.createElementNS(NS, "tspan");
        tspan.setAttribute("x", cx);
        tspan.setAttribute("y", startY + li * lineHeight);
        tspan.textContent = lines[li];
        textEl.appendChild(tspan);
      }

      fo.parentNode.replaceChild(textEl, fo);
    }

    const serializer = new XMLSerializer();
    return { ok: true, svg: serializer.serializeToString(svgEl) };
  });
}

// ─── Main ───────────────────────────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);
  const replaceFlag = args.includes("--replace");
  const mdPath = args.filter((a) => !a.startsWith("--"))[0];

  if (!mdPath) {
    console.error(
      "Usage: node convert_mermaid_to_excalidraw.mjs [--replace] <markdown-file>"
    );
    process.exit(1);
  }

  const resolvedMdPath = path.resolve(mdPath);
  if (!fs.existsSync(resolvedMdPath)) {
    console.error(`File not found: ${resolvedMdPath}`);
    process.exit(1);
  }

  let markdown = fs.readFileSync(resolvedMdPath, "utf-8");
  const blocks = extractMermaidBlocks(markdown);

  if (blocks.length === 0) {
    console.log("No ```mermaid code blocks found in the file.");
    process.exit(0);
  }

  console.log(`Found ${blocks.length} mermaid diagram(s). Converting...\n`);

  // Create output directory next to the markdown file
  const outDir = path.join(path.dirname(resolvedMdPath), "diagrams");
  fs.mkdirSync(outDir, { recursive: true });

  // Launch headless browser
  const browser = await chromium.launch();
  const page = await browser.newPage();

  // Load our converter page
  await page.setContent(buildConverterHTML());

  // Wait for the ESM modules to load
  try {
    await page.waitForFunction("window.__ready === true", { timeout: 60_000 });
  } catch {
    console.error(
      "ERROR: Timed out waiting for Mermaid / Excalidraw libraries to load.\n" +
        "Check your internet connection (CDN: esm.sh)."
    );
    await browser.close();
    process.exit(1);
  }

  const generated = [];

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const num = String(i + 1).padStart(2, "0");
    const slug = slugify(block.code, i);
    const baseName = `${num}-${slug}`;

    console.log(`[${num}] Converting: ${slug}`);

    // ── 1. Render Mermaid → SVG ──────────────────────────────────────
    const svgResult = await page.evaluate(
      async ({ code, id }) => window.renderMermaidSvg(code, id),
      { code: block.code, id: `mermaid-${i}` }
    );

    if (!svgResult.ok) {
      console.warn(`  ⚠  Mermaid SVG render failed: ${svgResult.error}`);
      console.warn(`     Skipping this block.`);
      generated.push(null);
      continue;
    }

    const svgPath = path.join(outDir, `${baseName}.svg`);

    // ── 1b. Post-process: replace foreignObject with native <text> ───
    const ppResult = await postProcessSvgInBrowser(page);
    let cleanSvg;
    if (ppResult.ok) {
      cleanSvg = ppResult.svg;
      console.log(`  ✓ Post-processed SVG (foreignObject → <text>)`);
    } else {
      console.warn(
        `  ⚠  Post-processing failed: ${ppResult.error}, using raw SVG`
      );
      cleanSvg = svgResult.svg;
    }

    fs.writeFileSync(svgPath, cleanSvg, "utf-8");
    console.log(
      `  ✓ SVG saved:            ${path.relative(process.cwd(), svgPath)}`
    );

    // ── 2. Convert Mermaid → Excalidraw elements ─────────────────────
    const excResult = await page.evaluate(
      async (code) => window.convertMermaidToExcalidraw(code),
      block.code
    );

    let excalidrawPath = null;

    if (excResult.ok) {
      // Convert skeleton elements to native Excalidraw format
      const nativeElements = convertSkeletonToNative(excResult.elements);
      const scene = buildExcalidrawScene(nativeElements, excResult.files);

      // Save .excalidraw (JSON — editable)
      excalidrawPath = path.join(outDir, `${baseName}.excalidraw`);
      fs.writeFileSync(
        excalidrawPath,
        JSON.stringify(scene, null, 2),
        "utf-8"
      );

      const textCount = nativeElements.filter(
        (e) => e.type === "text"
      ).length;
      const shapeCount = nativeElements.filter(
        (e) => e.type !== "text"
      ).length;
      console.log(
        `  ✓ Excalidraw saved:     ${path.relative(process.cwd(), excalidrawPath)} (${shapeCount} shapes + ${textCount} text elements)`
      );
    } else {
      console.warn(`  ⚠  Excalidraw conversion failed: ${excResult.error}`);
      console.warn(
        `     .excalidraw files will not be generated for this block.`
      );
    }

    generated.push({
      slug,
      baseName,
      svgPath,
      excalidrawPath,
      fullMatch: block.fullMatch,
    });

    console.log();
  }

  await browser.close();

  // ── 3. Optionally replace mermaid blocks in the markdown ────────────
  if (replaceFlag) {
    console.log("Replacing mermaid code blocks with image references...\n");

    for (let i = generated.length - 1; i >= 0; i--) {
      const g = generated[i];
      if (!g) continue;

      const imgFile = `diagrams/${g.baseName}.svg`;
      const altText = g.slug.replace(/-/g, " ");
      const imageRef = `![${altText}](${imgFile})`;

      markdown = markdown.replace(g.fullMatch, imageRef);
      console.log(`  [${String(i + 1).padStart(2, "0")}] → ${imageRef}`);
    }

    fs.writeFileSync(resolvedMdPath, markdown, "utf-8");
    console.log(`\n✓ Markdown updated: ${resolvedMdPath}`);
  } else {
    console.log(
      "Diagrams generated. Run with --replace to also update the markdown file."
    );
    console.log(
      "Example: node convert_mermaid_to_excalidraw.mjs --replace ../desigin.md"
    );
  }

  // ── Summary ─────────────────────────────────────────────────────────
  console.log("\n" + "═".repeat(60));
  console.log("Summary");
  console.log("═".repeat(60));
  const ok = generated.filter(Boolean).length;
  const skipped = generated.filter((g) => !g).length;
  console.log(`  Total blocks:     ${blocks.length}`);
  console.log(`  Converted:        ${ok}`);
  console.log(`  Skipped (errors): ${skipped}`);
  console.log(`  Output directory:  ${outDir}`);
  console.log("═".repeat(60));
}

main().catch((err) => {
  console.error("Fatal error:", err);
  process.exit(1);
});
