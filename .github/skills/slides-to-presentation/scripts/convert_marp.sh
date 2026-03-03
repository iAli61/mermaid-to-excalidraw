#!/usr/bin/env bash
#==============================================================================
# convert_marp.sh
#
# Wrapper script for Marp CLI exports: PDF, PPTX, HTML
# Handles SVG image path resolution and common options.
#
# Usage:
#   ./convert_marp.sh <slides.md> [--pdf] [--pptx] [--html] [--all]
#
# Prerequisites:
#   npm install -g @marp-team/marp-cli
#==============================================================================

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# ---------------------------------------------------------------------------
# Defaults
# ---------------------------------------------------------------------------
DO_PDF=false
DO_PPTX=false
DO_HTML=false

# ---------------------------------------------------------------------------
# Parse arguments
# ---------------------------------------------------------------------------
if [[ $# -lt 1 ]]; then
    echo "Usage: $0 <slides.md> [--pdf] [--pptx] [--html] [--all]"
    echo ""
    echo "Options:"
    echo "  --pdf    Export to PDF"
    echo "  --pptx   Export to PPTX (basic, no speaker notes)"
    echo "  --html   Export to HTML"
    echo "  --all    Export all formats"
    exit 1
fi

INPUT_FILE="$1"
shift

if [[ ! -f "$INPUT_FILE" ]]; then
    echo "ERROR: File not found: $INPUT_FILE"
    exit 1
fi

# Parse flags
while [[ $# -gt 0 ]]; do
    case "$1" in
        --pdf)  DO_PDF=true ;;
        --pptx) DO_PPTX=true ;;
        --html) DO_HTML=true ;;
        --all)  DO_PDF=true; DO_PPTX=true; DO_HTML=true ;;
        *)      echo "Unknown option: $1"; exit 1 ;;
    esac
    shift
done

# Default to PDF if no format specified
if ! $DO_PDF && ! $DO_PPTX && ! $DO_HTML; then
    DO_PDF=true
fi

# ---------------------------------------------------------------------------
# Check prerequisites
# ---------------------------------------------------------------------------
if ! command -v marp &> /dev/null; then
    echo "ERROR: Marp CLI not found."
    echo "Install with: npm install -g @marp-team/marp-cli"
    exit 1
fi

# ---------------------------------------------------------------------------
# Derive output paths
# ---------------------------------------------------------------------------
BASE_DIR="$(dirname "$INPUT_FILE")"
BASE_NAME="$(basename "$INPUT_FILE" .md)"

echo "════════════════════════════════════════════════════════════"
echo "  Marp Export"
echo "════════════════════════════════════════════════════════════"
echo "  Input:  $INPUT_FILE"
echo ""

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

if $DO_PDF; then
    OUTPUT="$BASE_DIR/$BASE_NAME.pdf"
    echo "  Exporting PDF..."
    marp --pdf --allow-local-files "$INPUT_FILE" -o "$OUTPUT" 2>&1
    echo "  ✓ PDF saved:  $OUTPUT"
fi

if $DO_PPTX; then
    OUTPUT="$BASE_DIR/$BASE_NAME.pptx"
    echo "  Exporting PPTX (basic, without speaker notes)..."
    marp --pptx --allow-local-files "$INPUT_FILE" -o "$OUTPUT" 2>&1
    echo "  ✓ PPTX saved: $OUTPUT"
    echo "  ⚠ Note: Marp PPTX does NOT include speaker notes."
    echo "    For PPTX with notes, use: python generate_pptx.py"
fi

if $DO_HTML; then
    OUTPUT="$BASE_DIR/$BASE_NAME.html"
    echo "  Exporting HTML..."
    marp --html --allow-local-files "$INPUT_FILE" -o "$OUTPUT" 2>&1
    echo "  ✓ HTML saved: $OUTPUT"
fi

echo ""
echo "════════════════════════════════════════════════════════════"
echo "  Done"
echo "════════════════════════════════════════════════════════════"
