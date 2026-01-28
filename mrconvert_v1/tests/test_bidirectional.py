from __future__ import annotations

import sys
import types
from pathlib import Path

import pytest

# Import the module we'll be testing
import mrconvert.bidirectional as conv


def _touch(p: Path) -> None:
    """Create a dummy file for testing"""
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(b"dummy content")


def test_pdf_to_docx_returns_conversion_result(tmp_path, monkeypatch):
    """Test PDF to DOCX conversion returns proper ConversionResult"""
    src = tmp_path / "test.pdf"
    dst = tmp_path / "test.docx"
    _touch(src)

    # Mock pdf2docx module
    class DummyConverter:
        def __init__(self, pdf_path: str) -> None:
            self.pdf_path = pdf_path

        def convert(self, docx_path: str) -> None:
            _touch(Path(docx_path))

        def close(self) -> None:
            pass

    def dummy_parse(pdf_path: str, docx_path: str) -> None:
        _touch(Path(docx_path))

    dummy_mod = types.SimpleNamespace(Converter=DummyConverter, parse=dummy_parse)
    monkeypatch.setitem(sys.modules, "pdf2docx", dummy_mod)

    # Test the conversion
    result = conv.pdf_to_docx(src, dst)

    assert result.input == src
    assert result.output == dst
    assert result.engine == "pdf2docx"
    assert dst.exists()


def test_docx_to_pdf_prefers_docx2pdf_on_windows(tmp_path, monkeypatch):
    """Test DOCX to PDF prefers docx2pdf on Windows/macOS"""
    src = tmp_path / "test.docx"
    dst = tmp_path / "test.pdf"
    _touch(src)

    # Mock docx2pdf module
    def fake_convert(input_path: str, output_path: str) -> None:
        _touch(Path(output_path))

    dummy_docx2pdf = types.SimpleNamespace(convert=fake_convert)
    monkeypatch.setitem(sys.modules, "docx2pdf", dummy_docx2pdf)

    # Mock platform to be Windows
    monkeypatch.setattr(conv.sys, "platform", "win32")

    result = conv.docx_to_pdf(src, dst)

    assert result.input == src
    assert result.output == dst
    assert result.engine == "docx2pdf"
    assert dst.exists()


def test_docx_to_pdf_fallback_to_soffice_on_linux(tmp_path, monkeypatch):
    """Test DOCX to PDF falls back to soffice on Linux when docx2pdf fails"""
    src = tmp_path / "test.docx"
    dst = tmp_path / "test.pdf"
    _touch(src)

    # Mock docx2pdf to raise exception (simulating Linux environment)
    def failing_convert(input_path: str, output_path: str) -> None:
        raise RuntimeError("docx2pdf not available on Linux")

    dummy_docx2pdf = types.SimpleNamespace(convert=failing_convert)
    monkeypatch.setitem(sys.modules, "docx2pdf", dummy_docx2pdf)

    # Mock platform to be Linux
    monkeypatch.setattr(conv.sys, "platform", "linux")

    # Mock shutil.which to return soffice path
    monkeypatch.setattr(
        conv.shutil,
        "which",
        lambda name: "/usr/bin/soffice" if name in ("soffice", "libreoffice") else None,
    )

    # Mock subprocess.run to simulate successful soffice execution
    def fake_run(cmd, **kwargs):
        # Simulate soffice creating the output file
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        produced = outdir / "test.pdf"
        _touch(produced)
        return types.SimpleNamespace(returncode=0, stdout="", stderr="")

    monkeypatch.setattr(conv.subprocess, "run", fake_run)

    result = conv.docx_to_pdf(src, dst)

    assert result.input == src
    assert result.output == dst
    assert result.engine == "soffice"
    assert dst.exists()


def test_docx_to_pdf_raises_error_when_no_engines_available(tmp_path, monkeypatch):
    """Test DOCX to PDF raises error when neither docx2pdf nor soffice available"""
    src = tmp_path / "test.docx"
    dst = tmp_path / "test.pdf"
    _touch(src)

    # Mock docx2pdf to fail
    def failing_convert(input_path: str, output_path: str) -> None:
        raise RuntimeError("docx2pdf not available")

    dummy_docx2pdf = types.SimpleNamespace(convert=failing_convert)
    monkeypatch.setitem(sys.modules, "docx2pdf", dummy_docx2pdf)

    # Mock shutil.which to return None (no soffice available)
    monkeypatch.setattr(conv.shutil, "which", lambda name: None)

    with pytest.raises(
        conv.EngineNotFoundError, match="docx2pdf unavailable and LibreOffice"
    ):
        conv.docx_to_pdf(src, dst)


def test_convert_file_auto_detects_pdf_input(tmp_path, monkeypatch):
    """Test convert_file auto-detects PDF input and converts to DOCX"""
    src = tmp_path / "test.pdf"
    _touch(src)

    # Mock pdf2docx
    class DummyConverter:
        def __init__(self, pdf_path: str) -> None:
            pass

        def convert(self, docx_path: str) -> None:
            _touch(Path(docx_path))

        def close(self) -> None:
            pass

    dummy_mod = types.SimpleNamespace(Converter=DummyConverter, parse=lambda *_: None)
    monkeypatch.setitem(sys.modules, "pdf2docx", dummy_mod)

    result = conv.convert_file(src)

    assert result.input == src
    assert result.output.suffix == ".docx"
    assert result.engine == "pdf2docx"


def test_convert_file_auto_detects_docx_input(tmp_path, monkeypatch):
    """Test convert_file auto-detects DOCX input and converts to PDF"""
    src = tmp_path / "test.docx"
    _touch(src)

    # Mock docx2pdf
    def fake_convert(input_path: str, output_path: str) -> None:
        _touch(Path(output_path))

    dummy_docx2pdf = types.SimpleNamespace(convert=fake_convert)
    monkeypatch.setitem(sys.modules, "docx2pdf", dummy_docx2pdf)
    monkeypatch.setattr(conv.sys, "platform", "win32")

    result = conv.convert_file(src)

    assert result.input == src
    assert result.output.suffix == ".pdf"
    assert result.engine == "docx2pdf"


def test_convert_file_raises_error_for_unsupported_format(tmp_path):
    """Test convert_file raises error for unsupported file formats"""
    src = tmp_path / "test.txt"
    _touch(src)

    with pytest.raises(conv.UnsupportedFormatError, match="Unsupported input type"):
        conv.convert_file(src)


def test_convert_file_raises_error_for_nonexistent_file(tmp_path):
    """Test convert_file raises error for non-existent files"""
    src = tmp_path / "nonexistent.pdf"

    with pytest.raises(FileNotFoundError):
        conv.convert_file(src)
