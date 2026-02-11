#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Tests for HWP Controller
"""

import pytest
import os
import tempfile
from unittest.mock import patch, MagicMock
from src.tools.hwp_controller import HwpController


class TestHwpController:
    """Test suite for HWP Controller."""

    @patch("win32com.client.Dispatch")
    @patch("win32com.client.GetActiveObject")
    def test_initialize_hwp(self, mock_get_active, mock_dispatch):
        """Test HWP initialization."""
        # Setup mock
        mock_hwp = MagicMock()
        mock_dispatch.return_value = mock_hwp
        mock_get_active.side_effect = Exception("Not running")

        # Initialize controller
        controller = HwpController()
        success = controller.connect(visible=False)

        # Verify initialization
        assert success is True
        assert controller.hwp is not None
        mock_dispatch.assert_called_with("HWPFrame.HwpObject")

    @patch("win32com.client.Dispatch")
    def test_open_document(self, mock_dispatch):
        """Test opening a document."""
        # Setup mock
        mock_hwp = MagicMock()
        mock_dispatch.return_value = mock_hwp
        mock_hwp.HAction.Execute.return_value = True

        # Initialize controller
        controller = HwpController()
        controller.hwp = mock_hwp
        controller.is_hwp_running = True

        # Test open document
        result = controller.open_document("test.hwp")

        # Verify results
        assert result is True
        assert controller.current_document_path.endswith("test.hwp")

    @patch("win32com.client.Dispatch")
    def test_save_document(self, mock_dispatch):
        """Test saving a document."""
        # Setup mock
        mock_hwp = MagicMock()
        mock_dispatch.return_value = mock_hwp

        # Initialize controller
        controller = HwpController()
        controller.hwp = mock_hwp
        controller.is_hwp_running = True

        # Test save document
        result = controller.save_document("test_saved.hwp")

        # Verify results
        assert result is True
        mock_hwp.SaveAs.assert_called()

    @patch("win32com.client.Dispatch")
    def test_get_text(self, mock_dispatch):
        """Test getting text from a document."""
        # Setup mock
        mock_hwp = MagicMock()
        mock_hwp.GetTextFile.return_value = "Test document content"
        mock_dispatch.return_value = mock_hwp

        # Initialize controller
        controller = HwpController()
        controller.hwp = mock_hwp
        controller.is_hwp_running = True

        # Test get text
        text = controller.get_text()

        # Verify results
        assert text == "Test document content"
        mock_hwp.GetTextFile.assert_called_with("TEXT", "")

    @patch("win32com.client.Dispatch")
    def test_insert_text(self, mock_dispatch):
        """Test inserting text into a document."""
        # Setup mock
        mock_hwp = MagicMock()
        mock_dispatch.return_value = mock_hwp

        # Initialize controller
        controller = HwpController()
        controller.hwp = mock_hwp
        controller.is_hwp_running = True

        # Test insert text
        result = controller.insert_text("Hello, World!")

        # Verify results
        assert result is True
        mock_hwp.HAction.Execute.assert_called()
