# AGENTS.md - HWP MCP Project Guide

This guide is for agentic coding agents working on the HWP MCP (Hangul Word Processor Model Context Protocol) project.

## Project Overview

This is a Python project that provides MCP server functionality for controlling Hangul Word Processor (HWP) through COM automation. The project includes:

- MCP server for HWP automation (`hwp_mcp_stdio_server.py`)
- HWP controller classes using win32com
- Table manipulation tools
- UI application using PySide6
- AI integration capabilities

## Build/Lint/Test Commands

### Environment Setup
```bash
# Create virtual environment (if not exists)
python -m venv .venv

# Activate virtual environment
# Windows:
.venv\Scripts\activate
# Unix:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Testing
```bash
# Run all tests
pytest

# Run tests with coverage
pytest --cov=src

# Run single test file
pytest src/__tests__/test_hwp_controller.py

# Run single test function
pytest src/__tests__/test_hwp_controller.py::TestHwpController::test_initialize_hwp

# Run tests with verbose output
pytest -v

# Run tests and stop on first failure
pytest -x
```

### Running the Application
```bash
# Run MCP server
python hwp_mcp_stdio_server.py

# Run UI application
python ui_app.py
```

## Code Style Guidelines

### General Python Conventions
- **Python Version**: Python 3.11+
- **Encoding**: UTF-8 (include `# -*- coding: utf-8 -*-` at top of files)
- **Indentation**: 4 spaces (no tabs)
- **Line Length**: Target under 100 characters, hard limit 120

### Imports
```python
# Standard library imports first
import os
import sys
from typing import Optional, List, Dict, Any, Tuple

# Third-party imports next
import win32com.client
import pytest
from PySide6.QtWidgets import QApplication, QWidget

# Local imports last
from src.tools.hwp_controller import HwpController
from .hwp_table_tools import HwpTableTools
```

### Type Hints
- Always use type hints for function parameters and return values
- Use `Optional[T]` for nullable types
- Use `Literal` for string enums
- Import from `typing` module

```python
from typing import Optional, List, Dict, Any, Tuple, Literal

def connect_document(path: str, visible: bool = True) -> Optional[HwpController]:
    """Connect to HWP document."""
    pass

Mode = Literal["rewrite", "summarize", "extend", "table"]
```

### Naming Conventions
- **Classes**: `PascalCase` (e.g., `HwpController`, `CommandParser`)
- **Functions/Variables**: `snake_case` (e.g., `get_hwp_controller`, `current_document_path`)
- **Constants**: `UPPER_SNAKE_CASE` (e.g., `AI_SERVER_REWRITE`)
- **Private members**: Prefix with underscore (e.g., `_current_hwp`, `_call_ai_server`)

### Docstrings
- Use triple quotes for docstrings
- Include Args, Returns, Raises sections where applicable
- Write docstrings in Korean for user-facing functions, English for internal functions

```python
def connect(self, visible: bool = True, register_security_module: bool = True) -> bool:
    """
    한글 프로그램에 연결합니다.

    Args:
        visible (bool): 한글 창을 화면에 표시할지 여부
        register_security_module (bool): 보안 모듈을 등록할지 여부

    Returns:
        bool: 연결 성공 여부
    """
    pass
```

### Error Handling
- Use specific exception types
- Log errors with appropriate level
- Return meaningful error messages to users
- Use try-catch blocks for COM operations

```python
try:
    self.hwp = win32com.client.GetActiveObject("HWPFrame.HwpObject")
    logger.info("GetActiveObject 성공 - 기존 HWP 인스턴스에 연결됨")
except Exception as e:
    logger.error(f"Failed to connect to HWP: {str(e)}", exc_info=True)
    return False
```

### Logging
- Use the `logging` module
- Create module-specific loggers
- Log at appropriate levels (DEBUG, INFO, WARNING, ERROR)
- Include exception info with `exc_info=True`

```python
import logging

logger = logging.getLogger("hwp-controller")

def some_function():
    logger.info("Starting operation")
    try:
        # operation
        logger.info("Operation completed successfully")
    except Exception as e:
        logger.error(f"Operation failed: {str(e)}", exc_info=True)
```

### File Organization
- **`src/tools/`**: Core HWP control classes
- **`src/utils/`**: Utility modules
- **`src/__tests__/`**: Test files
- **Root level**: Main entry points (server, UI app)

### MCP Tool Functions
- Use `@mcp.tool()` decorator for MCP server functions
- Include comprehensive docstrings with Args and Returns
- Handle exceptions gracefully and return error messages as strings
- Use JSON format for complex return data

```python
@mcp.tool()
def hwp_create_document_from_text(content: str, title: str = None) -> dict:
    """
    단일 문자열로 된 텍스트 내용으로 문서를 생성합니다.
    
    Args:
        content (str): 문서 내용
        title (str, optional): 문서 제목
        
    Returns:
        dict: 문서 생성 결과
    """
    try:
        # implementation
        return {"status": "success", "message": "Document created successfully"}
    except Exception as e:
        logger.error(f"Error creating document: {str(e)}", exc_info=True)
        return {"status": "error", "message": f"Error: {str(e)}"}
```

### Testing Guidelines
- Use `pytest` framework
- Mock COM objects with `unittest.mock`
- Test both success and failure cases
- Use descriptive test method names

```python
class TestHwpController:
    @patch('win32com.client.Dispatch')
    def test_initialize_hwp(self, mock_dispatch):
        """Test HWP initialization."""
        mock_hwp = MagicMock()
        mock_dispatch.return_value = mock_hwp
        
        controller = HwpController()
        
        mock_dispatch.assert_called_once_with("HWPFrame.HwpObject")
        assert controller.hwp is not None
```

### Security Considerations
- Validate file paths before operations
- Handle user input safely
- Don't log sensitive information
- Use proper error handling for COM operations

### COM Automation Best Practices
- Check if HWP is running before operations
- Handle connection losses gracefully
- Use proper COM object cleanup
- Test with different HWP versions

## Dependencies
- `pywin32>=228`: Windows COM automation
- `comtypes>=1.1.14`: Alternative COM library
- `pytest>=7.3.1`: Testing framework
- `pytest-cov>=4.1.0`: Coverage reporting
- `fastmcp>=0.1.0`: MCP server framework
- `PySide6`: GUI framework (for UI app)

## Common Patterns

### Connection Management
```python
def get_hwp_controller():
    """Get or create HwpController instance. Auto-reconnects if connection is lost."""
    global hwp_controller
    if hwp_controller is None:
        hwp_controller = HwpController()
        hwp_controller.connect(visible=True)
    return hwp_controller
```

### Error Response Format
```python
# For MCP tools that return strings
return f"Error: {error_message}"

# For functions returning dict
return {"status": "error", "message": error_message}
```

### Korean Text Handling
- Always use UTF-8 encoding
- Handle line breaks properly in HWP text insertion
- Use `ensure_ascii=False` for JSON operations with Korean text

## Development Notes
- This project specifically targets Korean HWP software
- COM automation requires HWP to be installed
- Test with actual HWP application for integration tests
- UI app uses PySide6 for cross-platform GUI development