---
name: design-to-presentation
description: >
  [DEPRECATED] This skill has been split into two separate skills:
  1. design-to-slides — creates slides.md, speaker-notes.md, and diagrams
  2. slides-to-presentation — exports to PDF, PPTX (python-pptx/VBA), HTML
  Use the two new skills instead.
---

# Design Document → Presentation Workshop Package

> **This skill has been split into two focused skills.**
> The scripts in `scripts/` remain here for backward compatibility.

## Skill 1: `design-to-slides`

**Location:** `.github/skills/design-to-slides/SKILL.md`

Transforms a technical design document (Markdown) into authoring artifacts:

| Output | Format | Purpose |
|--------|--------|---------|
| **Slide deck** | `.md` (Marp) | Source of truth — Markdown with Marp directives |
| **Speaker notes** | `.md` | Per-slide talking points, rationale, audience questions |
| **Diagrams** | `.svg` + `.excalidraw` | Visual assets — reused or newly created |

**Covers:** Document analysis, slide planning, Marp slide deck generation,
speaker notes generation, Mermaid → SVG/Excalidraw conversion.

## Skill 2: `slides-to-presentation`

**Location:** `.github/skills/slides-to-presentation/SKILL.md`

Converts existing slides and notes into distributable presentation formats:

| Output | Format | Speaker Notes? | Tool |
|--------|--------|----------------|------|
| **PDF handout** | `.pdf` | No | Marp CLI |
| **HTML slides** | `.html` | No | Marp CLI |
| **Full PPTX** | `.pptx` | **Yes** | python-pptx (`generate_pptx.py`) |
| **VBA PPTX** | `.pptx` | **Yes** | JSON export + VBA macro |

**Covers:** Marp CLI export, python-pptx generation with native tables and
speaker notes, SVG→PNG conversion, VBA macro workflow.

---

## Migration

Replace references to this skill:

| Old | New |
|-----|-----|
| `design-to-presentation` (creating slides) | `design-to-slides` |
| `design-to-presentation` (exporting PPTX/PDF) | `slides-to-presentation` |
