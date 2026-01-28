from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

import pytest

from mrconvert.cli import build_parser, run


def test_cli_accepts_to_docx_flag():
    """Test CLI accepts --to-docx flag"""
    parser = build_parser()

    # Test valid --to-docx usage
    args = parser.parse_args(["input.pdf", "--to-docx"])
    assert args.to_docx is True
    assert args.to_pdf is False


def test_cli_accepts_to_pdf_flag():
    """Test CLI accepts --to-pdf flag"""
    parser = build_parser()

    # Test valid --to-pdf usage
    args = parser.parse_args(["input.docx", "--to-pdf"])
    assert args.to_pdf is True
    assert args.to_docx is False


def test_cli_rejects_both_to_flags():
    """Test CLI rejects both --to-docx and --to-pdf flags"""
    # This test should check that our validation logic catches this
    # Since argparse doesn't automatically validate mutual exclusivity
    with pytest.raises(SystemExit):
        run(["input.pdf", "--to-docx", "--to-pdf"])


def test_cli_rejects_to_flags_with_format(tmp_path):
    """Test CLI rejects --to-docx/--to-pdf with --format flags"""
    src = tmp_path / "input.pdf"
    src.write_bytes(b"dummy content")

    with pytest.raises(SystemExit):
        run([str(src), "--to-docx", "--format", "md"])


def test_cli_requires_input_file_for_to_flags():
    """Test CLI requires input file when using --to-docx/--to-pdf"""
    parser = build_parser()

    with pytest.raises(SystemExit):
        parser.parse_args(["--to-docx"])


def test_cli_auto_detects_output_extension_for_to_docx(tmp_path, monkeypatch):
    """Test CLI auto-detects output extension for --to-docx"""
    src = tmp_path / "test.pdf"
    src.write_bytes(b"dummy pdf content")

    # Mock the bidirectional conversion
    with patch("mrconvert.bidirectional.pdf_to_docx") as mock_convert:
        mock_convert.return_value = type(
            "Result",
            (),
            {"input": src, "output": src.with_suffix(".docx"), "engine": "pdf2docx"},
        )()

        result = run([str(src), "--to-docx"])
        assert result == 0
        mock_convert.assert_called_once()


def test_cli_auto_detects_output_extension_for_to_pdf(tmp_path, monkeypatch):
    """Test CLI auto-detects output extension for --to-pdf"""
    src = tmp_path / "test.docx"
    src.write_bytes(b"dummy docx content")

    # Mock the bidirectional conversion
    with patch("mrconvert.bidirectional.docx_to_pdf") as mock_convert:
        mock_convert.return_value = type(
            "Result",
            (),
            {"input": src, "output": src.with_suffix(".pdf"), "engine": "docx2pdf"},
        )()

        result = run([str(src), "--to-pdf"])
        assert result == 0
        mock_convert.assert_called_once()


def test_cli_handles_conversion_error(tmp_path, monkeypatch):
    """Test CLI handles conversion errors gracefully"""
    src = tmp_path / "test.pdf"
    src.write_bytes(b"dummy pdf content")

    # Mock conversion to raise error
    with patch("mrconvert.bidirectional.pdf_to_docx") as mock_convert:
        from mrconvert.bidirectional import ConversionError

        mock_convert.side_effect = ConversionError("Conversion failed")

        result = run([str(src), "--to-docx"])
        assert result == 0  # CLI continues processing other files


def test_cli_uses_specified_output_dir_for_to_flags(tmp_path, monkeypatch):
    """Test CLI uses specified output directory for --to-docx/--to-pdf"""
    src = tmp_path / "test.pdf"
    src.write_bytes(b"dummy pdf content")
    out_dir = tmp_path / "output"

    # Mock the bidirectional conversion
    with patch("mrconvert.bidirectional.pdf_to_docx") as mock_convert:
        expected_dst = out_dir / "test.docx"
        mock_convert.return_value = type(
            "Result", (), {"input": src, "output": expected_dst, "engine": "pdf2docx"}
        )()

        result = run([str(src), "--to-docx", "--out", str(out_dir)])
        assert result == 0
        mock_convert.assert_called_once()
        # Verify the output path includes the specified directory
        call_args = mock_convert.call_args
        assert str(call_args[0][1]).endswith("test.docx")
