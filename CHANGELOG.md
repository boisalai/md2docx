# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.1] - 2025-10-31

### Changed
- Reorganized project structure with src/ layout
- Moved main code to src/md2docx/converter.py
- Created proper Python package structure with __init__.py
- Added CLI support via __main__.py (can now run with `python -m md2docx`)

### Fixed
- Updated README.md project structure documentation
- Fixed import paths in tests for new structure

## [1.1.0] - 2025-10-31

### Added
- Comprehensive input validation for DocumentConfig parameters
- Configuration constants for all magic numbers (font sizes, spacing, etc.)
- Logging system to replace print statements
- Verbose flag to control logging output
- Path traversal protection for security
- Markdown file extension validation
- Image format validation
- OS-specific Pandoc installation instructions
- Unit tests with pytest
- Type hints for all methods
- Support for development dependencies in pyproject.toml

### Changed
- Refactored code duplication: created `_set_language_for_run()` utility method
- Combined document post-processing and image insertion into single operation (performance improvement)
- Split `_apply_global_styles()` into smaller, focused methods
- Compiled regex patterns as class constants for better performance
- Improved error messages with more context
- Enhanced exception handling throughout
- Moved all imports to module level

### Fixed
- Image aspect ratio calculation bug (was using MAX_IMAGE_WIDTH instead of actual picture.width)
- Table border duplication issue
- Temporary file name collision handling
- Multiple document open/save operations reduced to single pass

### Removed
- Unused `extra_args` parameter from `convert()` method
- Redundant code for setting language properties

## [1.0.0] - 2025-10-31

### Added
- Initial release
- Markdown to Word conversion with Pandoc
- Custom heading styles and colors
- Image embedding with size adjustment
- Table formatting with borders
- Footnote support
- Custom headers and footers
- Language settings for spell-checking
- Multiple paper sizes (Letter, Legal, A4)
- Configurable document styles (Report, Note, Letter, Memo)
- Table of contents generation
- Professional formatting and styling
