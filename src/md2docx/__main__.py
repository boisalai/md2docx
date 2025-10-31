"""
Command-line interface for md2docx converter.

Allows running the converter as a module:
    python -m md2docx input.md output.docx
"""

import sys
import argparse
from pathlib import Path
from .converter import MarkdownToDocxConverter, DocumentConfig, PaperSize


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description='Convert Markdown documents to professionally formatted Word (DOCX) files.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic conversion
  python -m md2docx input.md output.docx

  # With custom configuration
  python -m md2docx input.md output.docx --author "John Doe" --language en-US

  # Quiet mode
  python -m md2docx input.md output.docx --quiet
"""
    )

    parser.add_argument('input', help='Input Markdown file')
    parser.add_argument('output', help='Output Word (DOCX) file')
    parser.add_argument('-w', '--working-dir', help='Working directory (default: current directory)')
    parser.add_argument('-a', '--author', default='', help='Document author')
    parser.add_argument('-d', '--date', default='', help='Document date')
    parser.add_argument('-l', '--language', default='en-US', help='Document language (default: en-US)')
    parser.add_argument('--paper-size', choices=['letter', 'legal', 'a4'], default='letter',
                       help='Paper size (default: letter)')
    parser.add_argument('--no-toc', action='store_true', help='Disable table of contents')
    parser.add_argument('-q', '--quiet', action='store_true', help='Quiet mode (only show errors)')
    parser.add_argument('--version', action='version', version='%(prog)s 1.1.3')

    args = parser.parse_args()

    # Create configuration
    paper_size_map = {
        'letter': PaperSize.LETTER,
        'legal': PaperSize.LEGAL,
        'a4': PaperSize.A4
    }

    config = DocumentConfig(
        author=args.author,
        date=args.date,
        language=args.language,
        paper_size=paper_size_map[args.paper_size],
        generate_toc=not args.no_toc
    )

    # Create converter
    converter = MarkdownToDocxConverter(config, verbose=not args.quiet)

    # Convert
    try:
        converter.convert(
            input_file=args.input,
            output_file=args.output,
            working_dir=args.working_dir
        )
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
