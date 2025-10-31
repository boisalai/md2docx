# md2docx

A Python tool for converting Markdown documents to professionally formatted Word (DOCX) files. The converter uses Pandoc for initial conversion and applies extensive post-processing to ensure consistent formatting, including custom styles, headers, footers, images, tables, and footnotes.

## Features

- Convert Markdown to Word with professional formatting
- Configurable document styles (Report, Note, Letter, Memo)
- Support for multiple paper sizes (Letter, Legal, A4)
- Custom heading colors and styles
- Automatic image embedding with size adjustment
- Table formatting with borders
- Footnote support with proper styling
- Custom headers and footers with page numbering
- Language settings for spell-checking
- Table of contents generation
- Configurable margins, fonts, and spacing
- Input validation and security (path traversal protection)
- Structured logging with configurable verbosity
- Comprehensive error handling

## Requirements

### System Dependencies

- Python 3.12 or higher
- Pandoc (external dependency for document conversion)

### Python Dependencies

- python-docx >= 1.2.0

## Installation

### 1. Install Pandoc

**macOS:**
```bash
brew install pandoc
```

**Linux (Debian/Ubuntu):**
```bash
sudo apt-get install pandoc
```

**Windows:**
Download and install from [https://pandoc.org/installing.html](https://pandoc.org/installing.html)

### 2. Install Python Package

```bash
# Clone the repository
git clone https://github.com/boisalai/md2docx.git
cd md2docx

# Install the package
pip install .

# Or install in development mode
pip install -e .

# Or with development dependencies
pip install -e ".[dev]"
```

## Usage

### Command Line Interface

The easiest way to use md2docx is via the command line:

```bash
# Basic conversion
python -m md2docx input.md output.docx

# With options
python -m md2docx input.md output.docx \
    --author "John Doe" \
    --date "2025-01-31" \
    --language en-US \
    --paper-size letter

# Quiet mode
python -m md2docx input.md output.docx --quiet

# See all options
python -m md2docx --help
```

### Python API - Basic Example

```python
from md2docx import MarkdownToDocxConverter, DocumentConfig

# Create a converter with default settings
converter = MarkdownToDocxConverter()

# Convert a markdown file
converter.convert(
    input_file="document.md",
    output_file="document.docx",
    working_dir="."
)
```

### Custom Configuration

```python
from md2docx import (
    MarkdownToDocxConverter,
    DocumentConfig,
    PaperSize,
    DocumentStyle
)

# Create custom configuration
config = DocumentConfig(
    style=DocumentStyle.REPORT,
    paper_size=PaperSize.LETTER,
    author="John Doe",
    date="2025-01-31",
    language="en-US",
    font_name="Arial",
    base_font_size=12,
    margins=(2.5, 2.5, 2.5, 2.5),  # top, right, bottom, left in cm
    line_spacing=1.15,
    generate_toc=True,
    heading_colors={
        1: (37, 150, 190),  # RGB color for Heading 1
        2: (37, 150, 190),  # RGB color for Heading 2
        3: (37, 150, 190)   # RGB color for Heading 3
    },
    footer_text={
        "odd": "Document Title",
        "even": "Author Name"
    }
)

# Create converter with custom config
converter = MarkdownToDocxConverter(config)

# Convert
converter.convert(
    input_file="report.md",
    output_file="report.docx"
)
```

### Using Preconfigured Styles

```python
from md2docx import MarkdownToDocxConverter, DocumentConfig, PaperSize

# Report style
config = DocumentConfig.create_report_style(
    author="Jane Smith",
    date="2025-01-31",
    language="en-US",
    generate_toc=True
)

# Note style
config = DocumentConfig.create_note_style(
    author="Jane Smith",
    date="2025-01-31",
    language="en-US"
)

converter = MarkdownToDocxConverter(config)
converter.convert("input.md", "output.docx")
```

## Configuration Options

### DocumentConfig Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `style` | DocumentStyle | REPORT | Document style template |
| `paper_size` | PaperSize | LETTER | Paper size (LETTER, LEGAL, A4) |
| `author` | str | "" | Document author |
| `date` | str | "" | Document date |
| `heading_colors` | Dict[int, Tuple[int, int, int]] | Blue | RGB colors for headings 1-3 |
| `footer_text` | Dict[str, str] | "Page" | Footer text for odd/even pages |
| `font_name` | str | "Arial" | Base font family |
| `base_font_size` | int | 12 | Base font size in points |
| `margins` | Tuple[float, float, float, float] | (2, 2, 2, 2) | Margins in cm (top, right, bottom, left) |
| `line_spacing` | float | 1.0 | Line spacing multiplier |
| `generate_toc` | bool | True | Generate table of contents |
| `language` | str | "en-US" | Document language code |

### Supported Paper Sizes

- `PaperSize.LETTER`: 8.5 x 11 inches (215.9 x 279.4 mm)
- `PaperSize.LEGAL`: 8.5 x 14 inches (215.9 x 355.6 mm)
- `PaperSize.A4`: 8.27 x 11.69 inches (210 x 297 mm)

### Document Styles

- `DocumentStyle.REPORT`: Professional report format
- `DocumentStyle.NOTE`: Internal note format
- `DocumentStyle.LETTER`: Letter format
- `DocumentStyle.MEMO`: Memo format

## Markdown Structure

### Heading Levels

The converter maps Markdown headings to Word styles as follows:

- `# Title` � Document Title (not a heading)
- `## Heading 1` � Heading 1 in Word
- `### Heading 2` � Heading 2 in Word
- `#### Heading 3` � Heading 3 in Word

### Images

Place images in an `img/` directory relative to your markdown file:

```markdown
![Image description](img/diagram.png)
```

Images will be automatically resized to fit the page width while maintaining aspect ratio.

### Tables

Standard Markdown tables are supported:

```markdown
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |
```

Tables will be formatted with borders and consistent styling.

### Footnotes

Pandoc-style footnotes are supported:

```markdown
Here is some text with a footnote.[^1]

[^1]: This is the footnote content.
```

## Project Structure

```
md2docx/
├── src/
│   └── md2docx/
│       ├── __init__.py       # Package initialization
│       ├── converter.py      # Main converter module
│       └── __main__.py       # CLI entry point
├── tests/
│   ├── __init__.py           # Tests package
│   └── test_md2docx.py       # Unit tests
├── README.md                 # Documentation
├── CHANGELOG.md              # Version history
├── LICENSE                   # License file
├── pyproject.toml            # Project configuration
├── .gitignore                # Git ignore rules
└── .python-version           # Python version specification
```

## Advanced Usage

### Custom Language Settings

Different languages can be specified for proper spell-checking:

```python
config = DocumentConfig(
    language="fr-CA"  # French (Canada)
)
```

Common language codes:
- `en-US`: English (United States)
- `en-GB`: English (United Kingdom)
- `fr-FR`: French (France)
- `fr-CA`: French (Canada)
- `es-ES`: Spanish (Spain)
- `de-DE`: German (Germany)

### Working with Multiple Documents

```python
from pathlib import Path
from md2docx import MarkdownToDocxConverter, DocumentConfig

# Process multiple files
config = DocumentConfig.create_report_style()
converter = MarkdownToDocxConverter(config)

input_dir = Path("./markdown_files")
output_dir = Path("./word_files")
output_dir.mkdir(exist_ok=True)

for md_file in input_dir.glob("*.md"):
    output_file = output_dir / f"{md_file.stem}.docx"
    converter.convert(
        input_file=md_file.name,
        output_file=str(output_file),
        working_dir=str(input_dir)
    )
```

### Logging and Verbosity Control

By default, the converter outputs informational messages during conversion. You can control this behavior with the `verbose` flag:

```python
from md2docx import MarkdownToDocxConverter, DocumentConfig

config = DocumentConfig.create_report_style()

# Quiet mode - only show warnings and errors
converter = MarkdownToDocxConverter(config, verbose=False)

# Verbose mode (default) - show all informational messages
converter = MarkdownToDocxConverter(config, verbose=True)

converter.convert("input.md", "output.docx")
```

The converter uses Python's standard `logging` module. You can configure it further:

```python
import logging

# Set custom logging level
logging.getLogger('md2docx').setLevel(logging.DEBUG)

# Or configure your own handler
handler = logging.FileHandler('conversion.log')
handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger('md2docx').addHandler(handler)
```

## Testing

### Running Tests

The project includes comprehensive unit tests. To run them:

```bash
# Install development dependencies
pip install -e ".[dev]"

# Run tests
pytest tests/ -v

# Run tests with coverage
pytest tests/ -v --cov=md2docx --cov-report=html

# Run a specific test file
pytest tests/test_md2docx.py -v

# Run a specific test
pytest tests/test_md2docx.py::TestDocumentConfig::test_default_config -v
```

### Test Coverage

Tests cover:
- Configuration validation
- Path handling and security
- Markdown parsing
- Image reference extraction
- Error handling
- Full integration tests (requires Pandoc)

## Troubleshooting

### Pandoc Not Found

If you receive an error about Pandoc not being installed:

1. Verify Pandoc is installed: `pandoc --version`
2. Ensure Pandoc is in your system PATH
3. Reinstall Pandoc following the installation instructions above

### Image Not Found

If images are not appearing in the output:

1. Verify images are in the `img/` directory
2. Check image file names match the Markdown references
3. Ensure image formats are supported (PNG, JPG, JPEG)

### Formatting Issues

If formatting doesn't appear correctly:

1. Verify your Markdown syntax is valid
2. Check that heading levels are consistent
3. Ensure custom configuration values are valid

## Contributing

Contributions are welcome. Please ensure that:

- Code follows the existing style and structure
- All functions include docstrings
- Changes are tested with various document types
- Commit messages are clear and descriptive

## License

MIT License - See LICENSE file for details

## Acknowledgments

This tool builds upon:
- Pandoc for initial document conversion
- python-docx for Word document manipulation
