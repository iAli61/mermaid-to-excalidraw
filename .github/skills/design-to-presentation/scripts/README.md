# Design Document → Presentation

Transforms a technical design document into a complete workshop delivery package.

## Quick Start

### 1. Generate Marp Slide Deck + Speaker Notes

Use the AI skill (Copilot) with the design document to generate:
- `workshop/slides.md` — Marp slide deck
- `workshop/speaker-notes.md` — Per-slide speaker notes

### 2. Convert Diagrams

```bash
cd scripts
node convert_mermaid_to_excalidraw.mjs --replace ../workshop/slides.md
```

### 3. Export PDF (Marp CLI)

```bash
# Install Marp CLI (one-time)
npm install -g @marp-team/marp-cli

# Export
bash .github/skills/design-to-presentation/scripts/convert_marp.sh workshop/slides.md --all
```

### 4. Generate PPTX with Speaker Notes (python-pptx)

```bash
# Install dependencies (one-time)
pip install -r .github/skills/design-to-presentation/scripts/requirements.txt

# Generate PPTX
python .github/skills/design-to-presentation/scripts/generate_pptx.py \
    --slides workshop/slides.md \
    --notes workshop/speaker-notes.md \
    --diagrams workshop/diagrams \
    --output workshop/presentation.pptx
```

### 5. Generate PPTX via VBA (Windows + PowerPoint)

```bash
# First, export JSON:
python .github/skills/design-to-presentation/scripts/generate_pptx.py \
    --slides workshop/slides.md \
    --notes workshop/speaker-notes.md \
    --export-json workshop/slides_data.json

# Then in PowerPoint:
# 1. Alt+F11 → File → Import → GeneratePresentation.bas
# 2. Alt+F8 → GeneratePresentation → Run
# 3. Select slides_data.json when prompted
```

## Output Files

| File | Format | Speaker Notes? | How Generated |
|------|--------|----------------|---------------|
| `slides.md` | Marp Markdown | No (separate file) | AI skill |
| `speaker-notes.md` | Markdown | Yes (the notes) | AI skill |
| `slides.pdf` | PDF | No | Marp CLI |
| `slides.pptx` | PPTX | No | Marp CLI |
| `presentation.pptx` | PPTX | **Yes** | python-pptx |
| `slides_data.json` | JSON | Yes (embedded) | generate_pptx.py |
| VBA-generated `.pptx` | PPTX | **Yes** | VBA macro |

## Scripts

| Script | Purpose |
|--------|---------|
| `generate_pptx.py` | Parse Marp + notes → PPTX with speaker notes (or export JSON) |
| `GeneratePresentation.bas` | VBA macro — reads JSON, creates PowerPoint with notes |
| `convert_marp.sh` | Wrapper for Marp CLI exports (PDF, PPTX, HTML) |
