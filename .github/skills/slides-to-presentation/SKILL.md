---
name: slides-to-presentation
description: >
  Convert Marp slide decks (slides.md) and speaker notes (speaker-notes.md)
  into presentation deliverables: PDF (via Marp CLI), PPTX with native tables
  and speaker notes (via python-pptx), HTML, or JSON + VBA macro for
  pixel-perfect PowerPoint. Handles SVG→PNG conversion for diagrams.
  Expects inputs from the design-to-slides skill.
---

# Slides & Notes → Presentation Exports

This skill converts **existing** Marp slide decks and speaker notes into
distributable presentation formats:

| Output | Format | Speaker Notes? | Tool |
|--------|--------|----------------|------|
| **PDF handout** | `.pdf` | No | Marp CLI |
| **HTML slides** | `.html` | No | Marp CLI |
| **Basic PPTX** | `.pptx` | No | Marp CLI |
| **Full PPTX** | `.pptx` | **Yes** | python-pptx (`generate_pptx.py`) |
| **VBA PPTX** | `.pptx` | **Yes** | JSON export + VBA macro |

> **Prerequisite:** The `slides.md`, `speaker-notes.md`, and `diagrams/`
> directory must already exist (created by the **`design-to-slides`** skill).

---

## Phase 1 — Export via Marp CLI (PDF / HTML / basic PPTX)

### 1.1 Prerequisites

Install Marp CLI globally:

```bash
npm install -g @marp-team/marp-cli
```

### 1.2 Export Commands

```bash
# PDF export (recommended for handouts)
marp --pdf --allow-local-files workshop/slides.md -o workshop/slides.pdf

# PPTX export (basic — no speaker notes)
marp --pptx --allow-local-files workshop/slides.md -o workshop/slides.pptx

# HTML export (interactive, keeps transitions)
marp --html --allow-local-files workshop/slides.md -o workshop/slides.html
```

> **Note:** Marp's PPTX export does NOT include speaker notes. For PPTX with
> speaker notes, use Phase 2 (python-pptx) or Phase 3 (VBA).

### 1.3 Wrapper Script

A convenience wrapper is provided:

```bash
bash .github/skills/slides-to-presentation/scripts/convert_marp.sh workshop/slides.md --all
```

Flags: `--pdf`, `--pptx`, `--html`, `--all`.

### 1.4 Image Path Fix

Marp resolves image paths relative to the Markdown file. Use
`--allow-local-files` to enable local file access. Ensure `diagrams/` is next
to the Markdown file.

---

## Phase 2 — Generate PPTX with Speaker Notes (python-pptx)

### 2.1 Prerequisites

```bash
pip install python-pptx Pillow cairosvg
```

Or use the requirements file:

```bash
pip install -r .github/skills/slides-to-presentation/scripts/requirements.txt
```

### 2.2 Script Location

```
.github/skills/slides-to-presentation/scripts/generate_pptx.py
```

### 2.3 Usage

```bash
python .github/skills/slides-to-presentation/scripts/generate_pptx.py \
  --slides workshop/slides.md \
  --notes workshop/speaker-notes.md \
  --diagrams workshop/diagrams \
  --output workshop/presentation.pptx
```

### 2.4 What It Does

1. **Parses** `slides.md` — splits on `---`, extracts titles, content, images, code blocks, tables
2. **Parses** `speaker-notes.md` — maps each `## Slide N` section to the corresponding slide
3. **Creates PPTX** with:
   - Title slides using the Title layout
   - Content slides with bullet points, bold formatting
   - Image slides with SVG→PNG conversion (via `cairosvg`)
   - Code slides with monospace formatting
   - **Native table slides** with styled header row, alternating shading
   - **Speaker notes** on every slide from the parsed notes
4. **Applies styling** — heading colors, font sizes matching the Marp theme

### 2.5 Table Rendering

Markdown tables in `slides.md` are rendered as **native PowerPoint table objects**:

| Feature | Detail |
|---------|--------|
| Header row | Blue background (`#2E5090`), white bold text |
| Data rows | Alternating light gray / white shading |
| Fonts | Calibri — 14pt header, 13pt body |
| Width | 85% of slide width, centered |
| Formatting | Markdown bold (`**`) and backticks stripped |

### 2.6 Export JSON for VBA (Optional)

```bash
python .github/skills/slides-to-presentation/scripts/generate_pptx.py \
  --slides workshop/slides.md \
  --notes workshop/speaker-notes.md \
  --export-json workshop/slides_data.json
```

This exports a structured JSON file that the VBA macro can consume.

---

## Phase 3 — Generate PowerPoint via VBA

### 3.1 When to Use VBA

Use VBA when:
- You need pixel-perfect PowerPoint formatting
- You want to use a corporate PowerPoint template (`.potx`)
- The python-pptx output needs manual refinement as a starting point
- You're on Windows with PowerPoint installed

### 3.2 Script Location

```
.github/skills/slides-to-presentation/scripts/GeneratePresentation.bas
```

### 3.3 Setup

1. **Generate the JSON data file** (Phase 2.6):
   ```bash
   python .github/skills/slides-to-presentation/scripts/generate_pptx.py \
     --slides workshop/slides.md \
     --notes workshop/speaker-notes.md \
     --export-json workshop/slides_data.json
   ```

2. **Open PowerPoint** on Windows

3. **Import the VBA macro:**
   - Press `Alt+F11` to open the VBA editor
   - Go to `File → Import File...`
   - Select `GeneratePresentation.bas`

4. **Run the macro:**
   - Press `Alt+F8`
   - Select `GeneratePresentation`
   - Click **Run**
   - When prompted, select the `slides_data.json` file

### 3.4 What the VBA Macro Does

1. **Reads** the JSON file containing slide data + speaker notes
2. **Creates** a new PowerPoint presentation
3. For each slide:
   - Selects the appropriate layout (Title, Content, Blank for images)
   - Populates the title and body text with formatting (bold, bullets)
   - Inserts images from the `diagrams/` directory (SVG or PNG)
   - Adds **speaker notes** from the notes data
4. **Applies** consistent font styling (Calibri, themed colors)
5. **Saves** the presentation

### 3.5 VBA Customization Points

| Setting | Location in VBA | Default |
|---|---|---|
| Slide width/height | `ActivePresentation.PageSetup` | Widescreen 16:9 |
| Title font & color | `TITLE_FONT_*` constants | Calibri 36pt #2E5090 |
| Body font & size | `BODY_FONT_*` constants | Calibri 20pt #333333 |
| Template file | `TEMPLATE_PATH` constant | (none — blank pres) |
| Image max dimensions | `IMG_MAX_*` constants | 9" × 5" |

---

## Phase 4 — Verification Checklist

### PPTX (python-pptx)
- [ ] All slides present with correct titles
- [ ] Speaker notes visible in Notes pane for every slide
- [ ] Tables render as native PPTX tables (not pipe-delimited text)
- [ ] Images render correctly (no broken references)
- [ ] Code blocks are readable with monospace font

### PDF (Marp CLI)
- [ ] All slides render without overflow
- [ ] Diagrams are crisp (not pixelated)
- [ ] Page count matches slide count

### HTML (Marp CLI)
- [ ] Slides advance correctly
- [ ] Images load from `diagrams/` directory

---

## Output Directory Structure

```
workshop/
├── slides.md                 ← Input (from design-to-slides skill)
├── speaker-notes.md          ← Input (from design-to-slides skill)
├── diagrams/                 ← Input (from design-to-slides skill)
│   ├── 01-architecture.svg
│   ├── 01-architecture.excalidraw
│   └── ...
├── slides.pdf                ← PDF export (Marp CLI)
├── slides.html               ← HTML export (Marp CLI)
├── presentation.pptx         ← PPTX with speaker notes (python-pptx)
├── slides_data.json          ← JSON intermediate (for VBA)
```

---

## Complete Workflow Summary

```
1. Ensure slides.md, speaker-notes.md, and diagrams/ exist
   └── Created by: design-to-slides skill
2. Export PDF via Marp CLI (Phase 1)
3. Export HTML via Marp CLI (Phase 1) — optional
4. Generate PPTX with speaker notes (Phase 2 — python-pptx)
   └── OR export JSON + run VBA macro (Phase 3)
5. Verify all outputs (Phase 4)
```

---

## Scripts Reference

| Script | Purpose | Location |
|--------|---------|----------|
| `generate_pptx.py` | Parse Marp + notes → PPTX with speaker notes (or export JSON) | `scripts/generate_pptx.py` |
| `GeneratePresentation.bas` | VBA macro — reads JSON, creates PowerPoint with notes | `scripts/GeneratePresentation.bas` |
| `convert_marp.sh` | Wrapper for Marp CLI exports (PDF, PPTX, HTML) | `scripts/convert_marp.sh` |
| `requirements.txt` | Python dependencies for `generate_pptx.py` | `scripts/requirements.txt` |
