<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# MD to DOCX Converter Project Instructions

This is a Python desktop application project that converts Markdown files to DOCX format using pandoc.

## Project Context
- **Language**: Python 3.7+
- **GUI Framework**: tkinter (standard library)
- **External Dependency**: pandoc (command-line tool)
- **Purpose**: Simple, user-friendly desktop application for file conversion

## Code Guidelines
- Follow PEP 8 Python style guidelines
- Use type hints where appropriate
- Include comprehensive error handling
- Maintain cross-platform compatibility (Windows, macOS, Linux)
- Keep the application as a single-file solution for easy distribution
- Use threading for non-blocking operations
- Provide clear user feedback and logging

## Architecture Principles
- **Single responsibility**: Each method should have one clear purpose
- **User experience**: Prioritize intuitive GUI and helpful error messages
- **Robustness**: Handle edge cases and provide graceful error recovery
- **Simplicity**: Keep the codebase maintainable and well-documented

## Key Features to Maintain
- File browser integration
- Real-time conversion progress
- Conversion options (TOC, numbering, templates)
- Comprehensive logging
- Pandoc integration and validation
