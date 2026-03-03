# Slides → Presentation Exports

Converts Marp slide decks and speaker notes into distributable formats
(PDF, PPTX with speaker notes, HTML).

## Prerequisites

- **Marp CLI** (for PDF/HTML/basic PPTX): `npm install -g @marp-team/marp-cli`
- **Python packages** (for PPTX with notes): `pip install -r requirements.txt`

## Quick Start

### 1. Export PDF / HTML (Marp CLI)

```bash
bash .github/skills/slides-to-presentation/scripts/convert_marp.sh workshop/slides.md --all
```

### 2. Generate PPTX with Speaker Notes (python-pptx)

```bash
python .github/skills/slides-to-presentation/scripts/generate_pptx.py \
    --slides workshop/slides.md \
    --notes workshop/speaker-notes.md \
    --diagrams workshop/diagrams \
    --output workshop/presentation.pptx
```

### 3. Generate PPTX via VBA (Windows + PowerPoint)

```bash
# First, export JSON:
python .github/skills/slides-to-presentation/scripts/generate_pptx.py \
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
| `slides.pdf` | PDF | No | Marp CLI |
| `slides.html` | HTML | No | Marp CLI |
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
| `requirements.txt` | Python dependencies for generate_pptx.py |
