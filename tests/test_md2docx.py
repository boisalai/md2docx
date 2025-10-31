"""
Unit tests for md2docx converter.

Run with: pytest tests/ -v
Or: python -m pytest tests/ -v
"""

import pytest
from pathlib import Path
import tempfile
import shutil

from md2docx import (
    MarkdownToDocxConverter,
    DocumentConfig,
    PaperSize,
    DocumentStyle
)


class TestDocumentConfig:
    """Tests for DocumentConfig class."""

    def test_default_config(self):
        """Test default configuration initialization."""
        config = DocumentConfig()
        assert config.style == DocumentStyle.REPORT
        assert config.paper_size == PaperSize.LETTER
        assert config.base_font_size == 12
        assert config.font_name == "Arial"
        assert config.language == "en-US"

    def test_custom_config(self):
        """Test custom configuration initialization."""
        config = DocumentConfig(
            style=DocumentStyle.NOTE,
            paper_size=PaperSize.A4,
            base_font_size=14,
            font_name="Times New Roman",
            language="fr-CA"
        )
        assert config.style == DocumentStyle.NOTE
        assert config.paper_size == PaperSize.A4
        assert config.base_font_size == 14
        assert config.font_name == "Times New Roman"
        assert config.language == "fr-CA"

    def test_report_style_factory(self):
        """Test report style factory method."""
        config = DocumentConfig.create_report_style(
            author="Test Author",
            date="2025-01-31"
        )
        assert config.style == DocumentStyle.REPORT
        assert config.author == "Test Author"
        assert config.date == "2025-01-31"

    def test_note_style_factory(self):
        """Test note style factory method."""
        config = DocumentConfig.create_note_style(
            author="Test Author",
            date="2025-01-31"
        )
        assert config.style == DocumentStyle.NOTE
        assert config.paper_size == PaperSize.LEGAL
        assert config.margins == (1.5, 1.5, 1.5, 1.5)

    def test_invalid_font_size(self):
        """Test validation of invalid font size."""
        with pytest.raises(ValueError, match="Base font size must be positive"):
            DocumentConfig(base_font_size=0)

        with pytest.raises(ValueError, match="Base font size must be positive"):
            DocumentConfig(base_font_size=-5)

    def test_invalid_margins(self):
        """Test validation of invalid margins."""
        with pytest.raises(ValueError, match="Margins must be non-negative"):
            DocumentConfig(margins=(2, -1, 2, 2))

    def test_invalid_line_spacing(self):
        """Test validation of invalid line spacing."""
        with pytest.raises(ValueError, match="Line spacing must be positive"):
            DocumentConfig(line_spacing=0)

    def test_invalid_heading_color(self):
        """Test validation of invalid heading color."""
        # Invalid RGB tuple (4 values)
        with pytest.raises(ValueError, match="must be RGB tuple"):
            DocumentConfig(heading_colors={1: (255, 255, 255, 0)})

        # Out of range RGB values
        with pytest.raises(ValueError, match="RGB values must be 0-255"):
            DocumentConfig(heading_colors={1: (300, 150, 100)})

    def test_invalid_footer_text(self):
        """Test validation of invalid footer text."""
        with pytest.raises(ValueError, match="footer_text must be a dict"):
            DocumentConfig(footer_text={"odd": "Test"})  # Missing 'even'


class TestMarkdownToDocxConverter:
    """Tests for MarkdownToDocxConverter class."""

    @pytest.fixture
    def temp_dir(self):
        """Create a temporary directory for tests."""
        temp_dir = tempfile.mkdtemp()
        yield Path(temp_dir)
        shutil.rmtree(temp_dir)

    @pytest.fixture
    def converter(self):
        """Create a converter instance for testing."""
        config = DocumentConfig()
        return MarkdownToDocxConverter(config, verbose=False)

    def test_converter_initialization(self):
        """Test converter initialization."""
        converter = MarkdownToDocxConverter()
        assert converter.config is not None
        assert converter.verbose is True

    def test_verbose_flag(self):
        """Test verbose flag functionality."""
        converter_verbose = MarkdownToDocxConverter(verbose=True)
        converter_quiet = MarkdownToDocxConverter(verbose=False)
        assert converter_verbose.verbose is True
        assert converter_quiet.verbose is False

    def test_extract_title_from_markdown(self, converter):
        """Test title extraction from markdown."""
        # Test with title
        content = "# My Document\n\nSome content here."
        title = converter._extract_title_from_markdown(content)
        assert title == "My Document"

        # Test without title
        content = "Some content without title."
        title = converter._extract_title_from_markdown(content)
        assert title == "Untitled Document"

        # Test with multiple headings
        content = "# First Title\n## Second Heading"
        title = converter._extract_title_from_markdown(content)
        assert title == "First Title"

    def test_extract_image_references(self, converter):
        """Test image reference extraction."""
        content = """
        # Document
        Some text.
        ![Image 1](img/image1.png)
        More text.
        ![Image 2](image2.jpg)
        """
        refs = converter._extract_image_references(content)
        assert len(refs) == 2
        assert refs[0]['alt_text'] == "Image 1"
        assert refs[0]['path'] == "image1.png"
        assert refs[1]['alt_text'] == "Image 2"
        assert refs[1]['path'] == "image2.jpg"

    def test_remove_image_references(self, converter):
        """Test image reference removal."""
        content = "Text before ![Image](img/test.png) text after"
        result = converter._remove_image_references(content)
        assert "[IMAGE_PLACEHOLDER]" in result
        assert "![Image]" not in result

    def test_setup_paths(self, converter, temp_dir):
        """Test path setup and validation."""
        work_dir, input_path, output_path = converter._setup_paths(
            "test.md",
            "test.docx",
            str(temp_dir)
        )
        # Use resolve() to handle symlinks (e.g., /var -> /private/var on macOS)
        assert work_dir.resolve() == temp_dir.resolve()
        assert input_path.resolve() == (temp_dir / "test.md").resolve()
        assert output_path.resolve() == (temp_dir / "test.docx").resolve()

    def test_setup_paths_invalid_directory(self, converter):
        """Test path setup with invalid directory."""
        with pytest.raises(FileNotFoundError, match="Working directory not found"):
            converter._setup_paths(
                "test.md",
                "test.docx",
                "/nonexistent/directory"
            )

    def test_setup_paths_traversal_protection(self, converter, temp_dir):
        """Test path traversal protection."""
        with pytest.raises(ValueError, match="within the working directory"):
            converter._setup_paths(
                "../outside.md",
                "test.docx",
                str(temp_dir)
            )

    def test_create_image_directory(self, converter, temp_dir):
        """Test image directory creation."""
        img_dir = converter._create_image_directory(temp_dir)
        assert img_dir.exists()
        assert img_dir == temp_dir / "img"

    def test_create_image_directory_already_exists(self, converter, temp_dir):
        """Test image directory when it already exists."""
        img_dir = temp_dir / "img"
        img_dir.mkdir()
        result = converter._create_image_directory(temp_dir)
        assert result == img_dir
        assert result.exists()

    def test_read_markdown_content(self, converter, temp_dir):
        """Test reading markdown content."""
        md_file = temp_dir / "test.md"
        content = "# Test\n\nThis is a test."
        md_file.write_text(content, encoding='utf-8')

        result = converter._read_markdown_content(md_file)
        assert result == content

    def test_read_markdown_content_nonexistent(self, converter, temp_dir):
        """Test reading nonexistent file."""
        md_file = temp_dir / "nonexistent.md"
        with pytest.raises(IOError, match="Cannot read markdown file"):
            converter._read_markdown_content(md_file)

    def test_create_temp_markdown(self, converter, temp_dir):
        """Test temporary markdown file creation."""
        input_path = temp_dir / "test.md"
        content = "# Test\n\n![Image](img/test.png)"

        temp_md = converter._create_temp_markdown(content, temp_dir, input_path)
        assert temp_md.exists()
        assert temp_md.name == "temp_test.md"

        temp_content = temp_md.read_text(encoding='utf-8')
        assert "[IMAGE_PLACEHOLDER]" in temp_content
        assert "![Image]" not in temp_content

    def test_cleanup_temp_markdown(self, converter, temp_dir):
        """Test temporary markdown cleanup."""
        temp_file = temp_dir / "temp_test.md"
        temp_file.write_text("test content")
        assert temp_file.exists()

        converter._cleanup_temp_markdown(temp_file)
        assert not temp_file.exists()

    def test_cleanup_nonexistent_temp_markdown(self, converter, temp_dir):
        """Test cleanup of nonexistent temp file (should not raise error)."""
        temp_file = temp_dir / "nonexistent.md"
        converter._cleanup_temp_markdown(temp_file)  # Should not raise

    def test_config_validation_constants(self):
        """Test that configuration constants are properly defined."""
        assert DocumentConfig.DEFAULT_TITLE_SIZE == 24
        assert DocumentConfig.DEFAULT_HEADING_1_SIZE == 18
        assert DocumentConfig.DEFAULT_HEADING_2_SIZE == 16
        assert DocumentConfig.DEFAULT_HEADING_3_SIZE == 14
        assert DocumentConfig.DEFAULT_TABLE_FONT_SIZE == 10
        assert DocumentConfig.DEFAULT_FOOTER_FONT_SIZE == 10
        assert DocumentConfig.DEFAULT_FOOTNOTE_FONT_SIZE == 10
        assert DocumentConfig.DEFAULT_PARAGRAPH_SPACING == 6
        assert DocumentConfig.DEFAULT_TABLE_CELL_SPACING == 2

    def test_converter_constants(self):
        """Test that converter constants are properly defined."""
        assert MarkdownToDocxConverter.SUPPORTED_IMAGE_EXTENSIONS == {
            '.png', '.jpg', '.jpeg', '.gif', '.bmp'
        }
        assert MarkdownToDocxConverter.IMAGE_PATTERN is not None


class TestIntegration:
    """Integration tests requiring pandoc."""

    @pytest.fixture
    def temp_dir(self):
        """Create a temporary directory for tests."""
        temp_dir = tempfile.mkdtemp()
        yield Path(temp_dir)
        shutil.rmtree(temp_dir)

    def test_full_conversion_simple(self, temp_dir):
        """Test full conversion of a simple markdown file."""
        # Check if pandoc is available
        if not shutil.which('pandoc'):
            pytest.skip("Pandoc not installed")

        # Create test markdown file
        md_content = """# Test Document

This is a test document with some **bold** and *italic* text.

## Section 1

Some content here.

### Subsection 1.1

More content.
"""
        md_file = temp_dir / "test.md"
        md_file.write_text(md_content, encoding='utf-8')

        # Create converter and convert
        config = DocumentConfig.create_report_style()
        converter = MarkdownToDocxConverter(config, verbose=False)

        converter.convert(
            input_file="test.md",
            output_file="test.docx",
            working_dir=str(temp_dir)
        )

        # Verify output file exists
        output_file = temp_dir / "test.docx"
        assert output_file.exists()
        assert output_file.stat().st_size > 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
