---
name: design-to-slides
description: >
  Transform a technical design document (Markdown) into a Marp slide deck
  (slides.md), detailed speaker notes (speaker-notes.md), and corresponding
  Mermaid diagrams. Reuses existing Mermaid diagrams or creates new ones,
  then converts them to SVG/Excalidraw via the mermaid-to-excalidraw skill.
  Does NOT handle export to PDF/PPTX — use the slides-to-presentation skill
  for that.
---

# Design Document → Slides & Speaker Notes

This skill transforms a technical design document (Markdown) into the
**authoring artifacts** of a workshop delivery package:

| Output | Format | Purpose |
|--------|--------|---------|
| **Slide deck** | `.md` (Marp) | Source of truth — Markdown with Marp directives |
| **Speaker notes** | `.md` | Per-slide talking points, rationale, audience questions |
| **Diagrams** | `.svg` + `.excalidraw` | Visual assets — reused or newly created |

> **Next step:** Once these files are ready, use the **`slides-to-presentation`**
> skill to export to PDF, PPTX (with speaker notes), or HTML.

---

## Phase 1 — Analyze the Design Document

### 1.1 Read & Identify Structure

Read the entire design document and map its sections to a slide outline:

| Document Pattern | Slide Type |
|---|---|
| Problem statement / motivation | Title + bullet slide |
| Architecture overview with environments | Diagram slide (flowchart) |
| Component anatomy or object models | Diagram slide (graph LR) |
| Step-by-step workflows or lifecycles | Diagram slide (flowchart TD) |
| Service interactions over time | Diagram slide (sequenceDiagram) |
| Decision trees / fallback chains | Diagram slide (flowchart TD, diamond nodes) |
| Tables (RBAC, config, comparison) | Table slide |
| Code snippets / pseudocode | Code slide |
| Implementation phases / roadmap | Numbered list slide |
| Open questions or TBDs | Discussion slide |

### 1.2 Plan Slide Flow

Structure slides in this order:

1. **Title slide** — project name, subtitle, date, pattern/approach
2. **Agenda** — numbered list of sections
3. **Problem Statement** — why this work matters
4. **Current State / Audit** — what exists today (findings, gaps)
5. **Architecture Overview** — high-level infrastructure + key decisions table
6. **Deep Dive(s)** — core concepts, catalogs, conventions
7. **Workflow / Process** — promotion, evaluation, CI/CD
8. **Integration** — how it connects to existing code
9. **Access Control** — RBAC, service principals
10. **Quality Gates** — evaluation criteria, metrics
11. **Operational** — rollback, monitoring, fallbacks
12. **Implementation Roadmap** — phased plan
13. **Open Questions** — items needing team input
14. **Q&A** — closing slide

### 1.3 Inventory Existing Visualizations

Before creating new diagrams, check if the design document already contains:

- **Mermaid code blocks** — reuse directly (copy into slides verbatim)
- **Existing `.svg` / `.excalidraw` files** in a `diagrams/` directory — reference with `![](diagrams/...)`
- **ASCII art or text diagrams** — candidates for conversion to Mermaid

**Decision rule:**
- If the document has a Mermaid block for a section → **reuse it** in the slide
- If the document has a section that needs visualization but no diagram → **create a new** Mermaid block
- If existing diagrams have already been converted to `.svg` → **reference** the SVG directly

---

## Phase 2 — Generate the Marp Slide Deck

### 2.1 Marp Front Matter

Every slide deck starts with Marp directives:

```yaml
---
marp: true
theme: gaia
paginate: true
backgroundColor: #fff
color: #333
style: |
  section {
    font-size: 28px;
  }
  section.lead h1 {
    font-size: 52px;
    color: #2e5090;
  }
  section.lead h2 {
    font-size: 32px;
    color: #555;
  }
  h1 {
    color: #2e5090;
  }
  h2 {
    color: #3a6bb5;
  }
  table {
    font-size: 22px;
  }
  code {
    font-size: 20px;
  }
  pre {
    font-size: 18px;
  }
  em {
    color: #d9534f;
    font-style: normal;
    font-weight: bold;
  }
  blockquote {
    border-left: 4px solid #2e5090;
    padding-left: 1rem;
    font-size: 24px;
    color: #555;
  }
---
```

### 2.2 Slide Content Rules

| Rule | Details |
|---|---|
| **One idea per slide** | Never combine two major concepts |
| **No walls of text** | Max 7 bullet points; max ~40 words per bullet |
| **Bold for impact** | Use `**bold**` for key terms and decisions |
| **Tables for comparisons** | Use Markdown tables instead of prose for anything tabular |
| **Code stays short** | Max ~15 lines of code per slide; use pseudocode with comments |
| **Quotes for key takeaways** | Use `>` blockquotes for one-sentence summaries |
| **Slide separators** | Use `---` on its own line between slides |
| **Lead class for section breaks** | Use `<!-- _class: lead -->` before title/transition slides |

### 2.3 Image Sizing for Marp

When referencing diagrams, **always** add Marp size directives to prevent
overflow. Use the `w:` and `h:` syntax in the alt text:

```markdown
![w:900 h:480](diagrams/01-architecture.svg)
```

**Guidelines by diagram type:**

| Diagram Type | Suggested Constraint | Rationale |
|---|---|---|
| Wide flowchart (LR) | `w:900 h:480` | Constrain width first |
| Tall flowchart (TD) with many nodes | `h:480` | Let width auto-scale |
| Sequence diagram | `w:900 h:450` | Usually wide and moderately tall |
| Object anatomy (graph LR) | `w:900 h:480` | Often very wide |

**Rule:** If the diagram's source viewBox height > 700px, always set `h:480`.
If width > 900px, always set `w:900`.

### 2.4 Handling Mermaid Diagrams in Slides

There are two approaches for diagrams in Marp slides:

**Option A — Keep Mermaid code blocks (for Marp CLI rendering):**
```markdown
```mermaid
flowchart TD
    A --> B
`` `
```

> Marp CLI renders Mermaid natively when using `--html` flag.

**Option B — Pre-render to SVG and reference (recommended):**
```markdown
![w:900 h:480](diagrams/01-architecture.svg)
```

> Pre-rendered SVGs are more reliable across viewers and allow Excalidraw editing.

**Recommendation:** Use Option B. First keep Mermaid blocks in the markdown for
authoring, then use the `mermaid-to-excalidraw` skill to convert and replace.

---

## Phase 3 — Generate Speaker Notes

### 3.1 Structure

Create a separate Markdown file (`speaker-notes.md`) with one section per slide:

```markdown
# Speaker Notes — [Presentation Title]

> **Companion document for `slides.md`**
> One section per slide.

---

## Slide 1 — [Slide Title]

### Talking Points
- Main point to communicate
- Supporting detail

### The "Why"
- Rationale behind the architectural choice shown on this slide
- Trade-offs considered

### Gotchas
- Common misconceptions or pushback to anticipate

---

## Slide 2 — [Slide Title]
...
```

### 3.2 Content Depth Rules

| Section | What to Include |
|---|---|
| **Talking Points** | What to say out loud. 3-5 bullets per slide. |
| **The "Why"** | Rationale behind design choices. Trade-offs. Alternatives considered. |
| **Gotchas** | Edge cases, common objections, things that break assumptions. |
| **Check for Understanding** | 1-2 audience questions (with suggested answers) per major section. |
| **Engagement Tips** | Suggested audience interactions (polls, show-of-hands). |

### 3.3 Quality Criteria

- Every slide in the deck has a corresponding section in speaker notes
- No slide is left without at least **Talking Points**
- At least **7-10 Check-for-Understanding questions** across the full deck
- **The "Why"** sections appear at least on architecture, design decision, and operational slides
- **Gotchas** appear on slides covering integration, promotion, and fallback topics
- Include a summary table of all Check-for-Understanding questions at the end

---

## Phase 4 — Convert Diagrams (Mermaid → SVG + Excalidraw)

### 4.1 Use the `mermaid-to-excalidraw` Skill

If the slides contain Mermaid code blocks:

```bash
cd scripts
node convert_mermaid_to_excalidraw.mjs ../workshop/slides.md
```

This generates `.svg` and `.excalidraw` files in `workshop/diagrams/`.

### 4.2 Replace and Size

```bash
node convert_mermaid_to_excalidraw.mjs --replace ../workshop/slides.md
```

Then **manually verify** each image reference has proper Marp size constraints
(see Phase 2.3).

### 4.3 Reuse Existing Diagrams

If the design document already has diagrams that were previously converted:

1. Check the existing `diagrams/` directory for `.svg` files
2. Reference them directly: `![w:900 h:480](diagrams/existing-diagram.svg)`
3. Only create new Mermaid blocks for sections that lack visualization

---

## Phase 5 — Verification Checklist

### Slide Deck (`slides.md`)
- [ ] Front matter has all Marp directives (theme, paginate, style)
- [ ] Slide count matches the document's logical sections
- [ ] No slide has more than 7 bullet points
- [ ] All diagram images have `w:` / `h:` size constraints
- [ ] Code blocks are ≤ 15 lines
- [ ] Lead class used for title and transition slides
- [ ] Tables have `font-size: 22px` applied via style

### Speaker Notes (`speaker-notes.md`)
- [ ] One `## Slide N` section per slide in the deck
- [ ] Every slide has at least **Talking Points**
- [ ] Architecture/decision slides have **The "Why"** sections
- [ ] Integration/operational slides have **Gotchas**
- [ ] 7-10 **Check for Understanding** questions total
- [ ] Summary table of questions at the end

### Diagrams
- [ ] All Mermaid blocks converted to `.svg` + `.excalidraw`
- [ ] SVG files have no `<foreignObject>` elements
- [ ] `.excalidraw` files have proper `containerId` / `boundElements`
- [ ] Image references in slides use Marp size constraints

---

## Output Directory Structure

```
workshop/
├── slides.md                 ← Marp slide deck (source of truth)
├── speaker-notes.md          ← Per-slide speaker notes
├── diagrams/
│   ├── 01-architecture.svg
│   ├── 01-architecture.excalidraw
│   ├── 02-workflow.svg
│   ├── 02-workflow.excalidraw
│   └── ...
```

> **Next step:** Use the **`slides-to-presentation`** skill to export these
> files to PDF, PPTX (with speaker notes), or HTML.

---

## Complete Workflow Summary

```
1. Read design document
2. Plan slide structure (Phase 1)
3. Inventory existing diagrams — reuse where possible
4. Write Marp slide deck (Phase 2)
5. Write speaker notes (Phase 3)
6. Convert Mermaid → SVG + Excalidraw (Phase 4)
   └── Uses: mermaid-to-excalidraw skill
7. Add image size constraints to slides
8. Verify all authoring outputs (Phase 5)
```
