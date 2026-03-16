---
name: notebook-to-description
description: >
  Analyze a Jupyter notebook (.ipynb) and generate a structured description
  document (Markdown) with Mermaid diagrams visualizing the notebook's
  workflow, architecture, and data flow. Converts diagrams to SVG and
  Excalidraw via the mermaid-to-excalidraw skill. Use when asked to describe,
  summarize, document, or visualize a notebook's structure and concepts.
---

# Notebook → Description with Visualizations

This skill analyzes a Jupyter notebook and produces a **structured description
document** with embedded diagrams that visualize the notebook's workflow,
architecture, and key concepts.

| Output | Format | Purpose |
|--------|--------|---------|
| **Description document** | `<notebook-stem>_description.md` | Structured summary with section headings, explanations, and diagram references |
| **Diagrams** | `.svg` + `.excalidraw` | Visual assets generated via the `mermaid-to-excalidraw` skill |

---

## Phase 1 — Analyze the Notebook

### 1.1 Read the Notebook

Read the entire notebook using the `copilot_getNotebookSummary` tool to get
cell structure, then read individual cells as needed. Collect:

- **Title and purpose** from the first markdown cell(s)
- **Section headings** from markdown cells throughout
- **Key imports and libraries** from code cells
- **Core functions, classes, and data structures** defined in code cells
- **Training loops, pipelines, or workflows** spanning multiple cells
- **Outputs and results** from markdown cells describing outcomes
- **Environment setup** (API keys, dependencies, hardware requirements)

### 1.2 Identify Diagram Candidates

Map notebook sections to diagram types:

| Notebook Pattern | Diagram Type |
|---|---|
| Multi-step pipeline (data → model → eval) | `flowchart TD` |
| Function call chains or class hierarchy | `flowchart LR` or `graph LR` |
| Agent-environment interaction loop | `flowchart TD` with cycle |
| Training loop with rollouts/episodes | `flowchart TD` with subgraphs |
| Data flow (input → transform → output) | `flowchart LR` |
| State machines or game logic | `flowchart TD` with diamond nodes |
| Temporal interactions (API calls, agent turns) | `sequenceDiagram` |
| Comparison of methods or models | `flowchart LR` with parallel subgraphs |
| Reward/loss computation breakdown | `flowchart TD` |
| Setup and dependency chain | `flowchart TD` |

### 1.3 Identify Key Concepts

For each major section, identify:

- **What** it does (functionality)
- **Why** it matters (purpose in the overall workflow)
- **How** it connects to other sections (dependencies, data flow)

---

## Phase 2 — Generate the Description Document

### 2.1 Document Structure

Create `<notebook-stem>_description.md` in the same directory as the notebook
with this structure:

```markdown
# <Notebook Title>

## Overview
<1-3 paragraph summary of what the notebook does, its purpose, and key outcomes>

## Prerequisites
<Libraries, API keys, hardware requirements, environment setup>

## Architecture
<High-level diagram of the overall workflow>

![architecture](diagrams/<nn>-<slug>.svg)

## Sections

### <Section N>: <Title>
<Description of what this section does and why>

![diagram](diagrams/<nn>-<slug>.svg)

<Key code constructs explained>

## Key Concepts
<Explanation of the core ML/AI/programming concepts used>

## Results and Outcomes
<What the notebook produces, metrics, comparisons>
```

### 2.2 Writing Guidelines

- **Be concrete**: Reference actual function names, variable names, and class
  names from the notebook using backticks.
- **Explain the "why"**: Don't just describe what code does — explain why each
  section exists in the overall workflow.
- **Connect sections**: Show how data flows from one section to the next.
- **Include parameter values**: Mention important hyperparameters, thresholds,
  and configuration values.
- **Use tables** for comparisons, parameter lists, or structured information.

### 2.3 Mermaid Diagram Guidelines

For each diagram:

- Use **fenced code blocks** with the `mermaid` language tag.
- Prefer `flowchart` over `graph` for directional layouts (TD, LR).
- Use `subgraph ID["Display Label"]` for grouped concepts.
- Keep node labels concise — use `<br>` for multi-line labels.
- Use meaningful edge labels to show data flow or relationships.
- Use styling to distinguish different types of nodes:
  - `fill:#e8f4fd,stroke:#2196F3` — input/setup nodes (blue)
  - `fill:#fff3e0,stroke:#FF9800` — processing/training nodes (orange)
  - `fill:#e8f5e9,stroke:#4CAF50` — output/result nodes (green)
  - `fill:#fce4ec,stroke:#E91E63` — alternative method nodes (pink)
  - `fill:#f3e5f5,stroke:#9C27B0` — shared/utility nodes (purple)
  - `fill:#ffebee,stroke:#f44336` — warning/error nodes (red)

---

## Phase 3 — Convert Diagrams to SVG and Excalidraw

After generating the description document with embedded Mermaid blocks,
use the **mermaid-to-excalidraw** skill to convert them:

### 3.1 Setup (if not already done)

```bash
cd ~/.copilot/skills/mermaid-to-excalidraw/scripts
npm install
npx playwright install chromium
```

### 3.2 Generate Diagrams

```bash
cd ~/.copilot/skills/mermaid-to-excalidraw/scripts
node convert_mermaid_to_excalidraw.mjs <path-to-description.md>
```

### 3.3 Replace Mermaid Blocks with Image References

```bash
cd ~/.copilot/skills/mermaid-to-excalidraw/scripts
node convert_mermaid_to_excalidraw.mjs --replace <path-to-description.md>
```

### 3.4 Output Files

| File | Purpose |
|---|---|
| `diagrams/NN-slug.svg` | Portable SVG — viewable everywhere |
| `diagrams/NN-slug.excalidraw` | Native Excalidraw JSON — editable in VS Code |

---

## Phase 4 — Verification

After generating the description and diagrams:

### Document Verification
- [ ] Description file exists next to the notebook
- [ ] All notebook sections are covered
- [ ] Function/class names match the actual notebook code
- [ ] Diagrams accurately represent the workflow
- [ ] No broken image references

### SVG Verification
- Confirm **zero** `<foreignObject>` elements (text must use native `<text>`)
- Confirm each SVG has `<text>` and `<tspan>` elements with readable content

### Excalidraw Verification
- Confirm **zero** elements with a `label` property (skeleton format eliminated)
- Confirm every text element has a `containerId` pointing to its parent shape
- Open a `.excalidraw` file in VS Code to visually confirm text is visible

---

## Complete Workflow Checklist

1. [ ] Read the notebook end to end (cells, structure, outputs)
2. [ ] Identify the overall purpose and key sections
3. [ ] Identify 3–8 sections that benefit from diagrams
4. [ ] Write the description document with Mermaid code blocks
5. [ ] Run `node convert_mermaid_to_excalidraw.mjs <description.md>` to generate `.svg` + `.excalidraw`
6. [ ] Verify SVG files: no `<foreignObject>`, proper `<text>` elements
7. [ ] Verify `.excalidraw` files: no `label` properties, proper `containerId`/`boundElements`
8. [ ] Run with `--replace` to swap Mermaid blocks for `![](diagrams/...)` image refs
9. [ ] Commit `diagrams/` directory alongside the description document

---

## Example Invocation

Given a notebook `session_01_frozen_lake.ipynb`, this skill produces:

```
session_01_frozen_lake_description.md    # Structured description
diagrams/
  01-overall-architecture.svg            # High-level workflow
  01-overall-architecture.excalidraw
  02-environment-logic.svg               # Game mechanics
  02-environment-logic.excalidraw
  03-training-loop.svg                   # RL training pipeline
  03-training-loop.excalidraw
  04-model-comparison.svg                # Benchmark results
  04-model-comparison.excalidraw
```
