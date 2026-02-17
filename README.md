# Document to Markdown Converter

A powerful Python tool for converting PDF, PPTX, and DOCX files into high-fidelity Markdown with AI-powered text extraction, image deduplication, and parallel processing.

![Python](https://img.shields.io/badge/python-3.13-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## Features

### 🎯 Core Capabilities
- **Multi-format Support**: Convert PDF, PPTX, and DOCX to Markdown
- **AI-Powered Extraction**: Uses `marker-pdf` for intelligent text and layout recognition
- **Faithful Copy**: Generates JPEG snapshots of each page/slide for visual fidelity
- **Image Deduplication**: SHA-256 hashing prevents duplicate images (logos, icons)
- **Metadata Extraction**: Automatic YAML frontmatter with title, author, date, and page count

### ⚡ Performance
- **Parallel Processing**: Multi-core batch conversion with `ProcessPoolExecutor`
- **Optimized Snapshots**: JPEG compression (80% quality, 1.5x zoom) achieves >60% size reduction
- **Single-Pass Interleaving**: O(N) regex substitution for snapshot placement
- **CPU Fallback**: Automatic retry on CPU if GPU (MPS) extraction fails

### 🖥️ Mac Optimization
- **MPS Support**: Metal Performance Shaders acceleration on Apple Silicon
- **Graceful Interrupts**: Clean Ctrl+C handling without stack trace spam
- **Resource Management**: Conservative worker defaults to prevent memory exhaustion

## Installation

### Prerequisites
- Python 3.13+
- macOS (for PPTX snapshot generation via PowerPoint)
- Microsoft PowerPoint (optional, for slide snapshots)

### Setup
```bash
# Clone the repository
git clone https://github.com/phyrexia/document-to-markdown.git
cd document-to-markdown

# Install dependencies
pip install -r requirements.txt

# Optional: Install certifi for SSL support
pip install certifi
```

## Usage

### Basic Conversion
```bash
# Single file
python Transformer_v3_Parallel.py document.pdf

# Multiple files
python Transformer_v3_Parallel.py file1.pdf file2.pptx file3.docx

# With shell script wrapper
./run_transformer.sh document.pdf
```

### Advanced Options
```bash
# Skip snapshots for faster conversion (no PowerPoint automation)
python Transformer_v3_Parallel.py --no-snapshots presentation.pptx

# Limit pages for testing
python Transformer_v3_Parallel.py --pages 10 large_document.pdf

# Control parallelism
python Transformer_v3_Parallel.py --workers 4 *.pdf

# Force CPU mode
python Transformer_v3_Parallel.py --device cpu document.pdf

# Enable debug logging
python Transformer_v3_Parallel.py --debug document.pdf
```

## Versions

| Version | Description | Use Case |
|---------|-------------|----------|
| `Transformer_v3_Parallel.py` | **Latest**. Parallel processing with `--no-snapshots` support | Batch conversion, production use |
| `Transformer_v2_Optimized.py` | Stable single-threaded with metadata and optimizations | Single files, guaranteed stability |
| `Transformer.py` | Legacy version | Not recommended |

## Output Format

### Markdown Structure
```markdown
---
title: Document Name
processed_at: 2026-02-04 14:00:00
original_author: John Doe
original_title: Monthly Report
page_count: 12
---

![Page 1 Preview](document_images/page_1_snapshot.jpg)

<span id="page-1-0"></span>

# Document Title

Extracted text content...

![Extracted Image](document_images/image_1.jpg)
```

### Output Files
- `document.md`: Markdown file with frontmatter
- `document_images/`: Directory containing:
  - `page_N_snapshot.jpg`: Full page/slide previews
  - `image_N.jpg`: Extracted embedded images (deduplicated)

## Architecture

### Key Components
- **DocumentConverter**: Main class handling all conversion logic
- **Snapshot Generation**: Uses `fitz` (PyMuPDF) for PDF rendering
- **AI Extraction**: `marker-pdf` with `surya-ocr` for layout analysis
- **Parallel Execution**: `ProcessPoolExecutor` with spawn method for MPS safety

### Image Processing Pipeline
1. **Snapshot Generation**: Render pages/slides as JPEG (1.5x zoom, 80% quality)
2. **AI Extraction**: Extract embedded images via `marker-pdf`
3. **Deduplication**: SHA-256 hash comparison
4. **Interleaving**: Single-pass regex substitution to insert snapshots

## Performance Benchmarks

| Document Type | Pages | Time (v2) | Time (v3, 4 workers) | Speedup |
|---------------|-------|-----------|----------------------|---------|
| PDF (Text-heavy) | 50 | 45s | 15s | 3x |
| PPTX (with snapshots) | 20 | 60s | 20s | 3x |
| PPTX (--no-snapshots) | 20 | 20s | 7s | 2.8x |

## Troubleshooting

### Common Issues

**PowerPoint opens during PPTX conversion**
- This is expected for snapshot generation
- Use `--no-snapshots` to skip (faster, but no slide previews)

**"Recognizing Layout" hangs**
- Reduce workers: `--workers 2`
- Force CPU: `--device cpu`
- Skip snapshots: `--no-snapshots`

**KeyboardInterrupt stack traces**
- Fixed in v3 with graceful signal handling
- Update to latest version

**Missing images in PDF**
- CPU fallback should handle this automatically
- Check logs for "Primary extraction failed... Retrying on CPU"

## Development

### Project Structure
```
.
├── Transformer_v3_Parallel.py    # Latest parallel version
├── Transformer_v2_Optimized.py   # Stable single-thread
├── requirements.txt               # Python dependencies
├── run_transformer.sh             # Shell wrapper
└── README.md                      # This file
```

### Contributing
Contributions are welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request with clear description

## License

MIT License - See LICENSE file for details

## Acknowledgments

- [marker-pdf](https://github.com/VikParuchuri/marker) - AI-powered PDF extraction
- [surya-ocr](https://github.com/VikParuchuri/surya) - Layout analysis and OCR
- [PyMuPDF](https://pymupdf.readthedocs.io/) - PDF rendering
- [python-pptx](https://python-pptx.readthedocs.io/) - PPTX parsing
- [python-docx](https://python-docx.readthedocs.io/) - DOCX parsing

## Author

**Gonzalo González Pineda**
- GitHub: [@phyrexia](https://github.com/phyrexia)

---

**Note**: This tool was developed for personal use and optimized for macOS with Apple Silicon. Windows/Linux support may require modifications to the PPTX snapshot generation logic.
