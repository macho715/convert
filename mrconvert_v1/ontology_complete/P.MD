---
title: "P.MD - PDF↔DOCX Bidirectional Converter"
type: "technical-specification"
domain: "document-conversion"
version: "1.0"
date: "2025-01-19"
tags: ["pdf", "docx", "converter", "bidirectional", "python", "mrconvert"]
libraries: ["pdf2docx", "docx2pdf", "libreoffice"]
status: "active"
---

---
title: "PDF↔DOCX 양방향 변환기"
type: "tool-documentation"
domain: "file-conversion"
version: "1.0"
date: "2025-01-19"
tags: ["pdf", "docx", "conversion", "bidirectional", "python"]
status: "active"
---

# P.MD - PDF↔DOCX 양방향 변환기

## 개요

"PDF나 Word(DOCX)를 올리면 반대로 변환해주는" 양방향 변환기를 바로 쓸 수 있는 파이썬 프로그램으로 묶어 드립니다. 기본은 PDF → DOCX(pdf2docx), DOCX → PDF(docx2pdf 또는 LibreOffice soffice 폴백)입니다.

docx2pdf는 Windows/macOS에서 Microsoft Word가 설치돼 있어야 동작합니다. 리눅스에선 지원하지 않으니 LibreOffice로 자동 폴백시키겠습니다.

**GitHub**: [pdf2docx](https://github.com/py-pdf/pypdf) | [docx2pdf](https://github.com/AlJohri/docx2pdf)
**PyPI**: [pdf2docx](https://pypi.org/project/pdf2docx/) | [docx2pdf](https://pypi.org/project/docx2pdf/)

pdf2docx는 내부적으로 PyMuPDF로 PDF를 분석하고 python-docx로 DOCX를 생성합니다. 예제 사용법/설명은 아래 레퍼런스에 따라 짰습니다.

**GitHub**: [pdf2docx](https://github.com/py-pdf/pypdf)
+2
pdf2docx.readthedocs.io
+2

LibreOffice 폴백은 표준 CLI(soffice --headless --convert-to pdf …)로 구현합니다.
Ask LibreOffice

1) 프로젝트 구조
pdf-word-converter/
├─ pyproject.toml
├─ README.md
├─ src/
│  └─ pdf_word_converter/
│     ├─ __init__.py
│     ├─ converter.py      # 변환 로직 (핵심)
│     └─ cli.py            # 터미널용 CLI
├─ web/
│  └─ app.py               # (선택) 업로드용 초간단 FastAPI
└─ tests/
   └─ test_converter.py    # 빠른 단위 테스트 (I/O 더블)

2) 설치
# 가상환경 권장
pip install -U pdf2docx docx2pdf fastapi uvicorn
# (리눅스/서버에서 DOCX→PDF가 필요하면 LibreOffice 설치 필수)
# ex) Debian/Ubuntu: sudo apt-get install -y libreoffice


참고: docx2pdf는 Word가 있는 Windows/macOS에서만 동작합니다. 리눅스/서버에선 자동으로 soffice 폴백을 씁니다.
GitHub
+2
PyPI
+2

3) 핵심 로직 (src/pdf_word_converter/converter.py)
# pyright: reportMissingImports=false
from __future__ import annotations

import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional


@dataclass(frozen=True)
class ConversionResult:
    input: Path
    output: Path
    engine: str  # 'pdf2docx' | 'docx2pdf' | 'soffice'


class ConversionError(Exception):
    pass


def convert_file(
    in_path: str | Path,
    out_path: str | Path | None = None,
    *,
    target: Optional[str] = None,
) -> ConversionResult:
    """PDF<->DOCX 단방향 자동 판별 변환기."""
    src = Path(in_path).expanduser().resolve()
    if not src.exists():
        raise FileNotFoundError(src)

    suffix = src.suffix.lower()
    if target:
        target = target.lower().lstrip(".")

    match suffix:
        case ".pdf":
            dst = Path(out_path) if out_path else src.with_suffix(".docx")
            return _pdf_to_docx(src, dst)
        case ".docx":
            dst = Path(out_path) if out_path else src.with_suffix(".pdf")
            return _docx_to_pdf(src, dst)
        case _:
            raise ConversionError(f"Unsupported input type: {suffix} (use .pdf or .docx)")


def _pdf_to_docx(src: Path, dst: Path) -> ConversionResult:
    # pdf2docx는 PyMuPDF로 PDF를 분석하고 python-docx로 DOCX 생성
    # ref: https://github.com/ArtifexSoftware/pdf2docx , docs quickstart
    from pdf2docx import Converter, parse  # type: ignore

    dst.parent.mkdir(parents=True, exist_ok=True)
    try:
        cv = Converter(str(src))
        try:
            cv.convert(str(dst))  # 모든 페이지
        finally:
            cv.close()
    except Exception:
        # 보조 경로: parse() 한 방
        parse(str(src), str(dst))

    if not dst.exists():
        raise ConversionError(f"pdf2docx failed to create {dst}")
    return ConversionResult(src, dst, "pdf2docx")


def _docx_to_pdf(src: Path, dst: Path) -> ConversionResult:
    # 1) Word가 있는 Windows/macOS면 docx2pdf 우선
    try:
        from docx2pdf import convert as d2p_convert  # type: ignore

        if sys.platform in {"win32", "darwin"}:
            dst.parent.mkdir(parents=True, exist_ok=True)
            d2p_convert(str(src), str(dst))
            if dst.exists():
                return ConversionResult(src, dst, "docx2pdf")
    except Exception:
        # Linux 등 환경에선 예외가 날 수 있음 → 폴백
        pass

    # 2) 리눅스/서버 폴백: LibreOffice 'soffice'
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise ConversionError(
            "docx2pdf unavailable and LibreOffice (soffice) not found in PATH."
        )

    outdir = dst.parent
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(outdir), str(src)]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if proc.returncode != 0:
        raise ConversionError(f"LibreOffice failed: {proc.stderr or proc.stdout}")

    produced = outdir / (src.with_suffix(".pdf").name)
    if produced != dst and produced.exists():
        produced.rename(dst)

    if not dst.exists():
        raise ConversionError(f"Expected {dst} but not found after conversion.")
    return ConversionResult(src, dst, "soffice")


pdf2docx 동작/예제는 공식 문서의 Converter(...).convert(...)/parse(...) 사용법을 따릅니다.
pdf2docx.readthedocs.io

docx2pdf는 “Windows/macOS의 Word를 호출”하는 방식이라 Linux 미지원입니다(이슈/문서에 명시). 그래서 soffice 폴백을 넣었습니다.
GitHub
+2
PyPI
+2

soffice CLI는 실무에서 널리 쓰이는 표준 호출입니다.
Ask LibreOffice

4) CLI (src/pdf_word_converter/cli.py)
from __future__ import annotations

import argparse
import json
from pathlib import Path

from .converter import ConversionError, convert_file


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        prog="pdf-word-converter",
        description="Convert PDF<->DOCX by auto-detecting input type.",
    )
    p.add_argument("inputs", nargs="+", help="Input file path(s): .pdf or .docx")
    p.add_argument("-o", "--outdir", help="Output directory (optional)")
    p.add_argument("--json", action="store_true", help="Print JSON result")
    args = p.parse_args(argv)

    results = []
    for ip in args.inputs:
        src = Path(ip)
        if args.outdir:
            suffix = ".pdf" if src.suffix.lower() == ".docx" else ".docx"
            dst = Path(args.outdir) / (src.stem + suffix)
        else:
            dst = None

        try:
            res = convert_file(src, dst)
            results.append({"input": str(res.input), "output": str(res.output), "engine": res.engine})
            if not args.json:
                print(f"[{res.engine}] {res.input.name} → {res.output}")
        except Exception as e:
            if not args.json:
                print(f"[ERROR] {src}: {e}")
            results.append({"input": str(src), "error": str(e)})

    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


사용 예:

# PDF → DOCX
python -m pdf_word_converter.cli sample.pdf

# DOCX → PDF (여러 파일 한 번에, 아웃풋 디렉터리 지정)
python -m pdf_word_converter.cli a.docx b.docx -o ./out

5) (선택) 웹 업로더 (web/app.py)
from __future__ import annotations

import tempfile
from pathlib import Path

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse

from src.pdf_word_converter.converter import convert_file, ConversionError

app = FastAPI()


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    suffix = Path(file.filename).suffix.lower()
    if suffix not in {".pdf", ".docx"}:
        return JSONResponse({"error": "Only .pdf or .docx accepted"}, status_code=400)

    with tempfile.TemporaryDirectory() as td:
        tmp_in = Path(td) / file.filename
        tmp_out = tmp_in.with_suffix(".docx" if suffix == ".pdf" else ".pdf")
        data = await file.read()
        tmp_in.write_bytes(data)

        try:
            res = convert_file(tmp_in, tmp_out)
            return FileResponse(
                path=res.output,
                filename=Path(file.filename).with_suffix(tmp_out.suffix).name,
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                if tmp_out.suffix == ".docx"
                else "application/pdf",
            )
        except ConversionError as e:
            return JSONResponse({"error": str(e)}, status_code=500)


실행:

uvicorn web.app:app --reload
# POST /convert 로 파일 업로드하면 변환된 파일이 바로 다운로드됩니다.

6) 빠른 단위 테스트 (tests/test_converter.py)

실제 변환 라이브러리를 때리지 않고 더블/주입으로 빠르게 검증합니다(≤200ms 목표).

from __future__ import annotations

import sys
from pathlib import Path

import builtins
import types

import pdf_word_converter.converter as conv


def _touch(p: Path):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_bytes(b"x")


def test_pdf_to_docx_monkeypatch(tmp_path, monkeypatch):
    src = tmp_path / "x.pdf"
    _touch(src)

    class DummyCv:
        def __init__(self, *_): ...
        def convert(self, out): _touch(Path(out))
        def close(self): ...

    dummy_mod = types.SimpleNamespace(Converter=DummyCv, parse=lambda *_: None)
    monkeypatch.setitem(sys.modules, "pdf2docx", dummy_mod)

    res = conv.convert_file(src)
    assert res.output.suffix == ".docx"
    assert res.engine == "pdf2docx"
    assert res.output.exists()


def test_docx_to_pdf_prefers_docx2pdf_on_macos(tmp_path, monkeypatch):
    src = tmp_path / "x.docx"
    _touch(src)
    dst = src.with_suffix(".pdf")

    def fake_convert(i, o):
        _touch(Path(o))

    dummy = types.SimpleNamespace(convert=fake_convert)
    monkeypatch.setitem(sys.modules, "docx2pdf", dummy)
    monkeypatch.setattr(conv.sys, "platform", "darwin")

    res = conv.convert_file(src)
    assert res.engine == "docx2pdf"
    assert dst.exists()


def test_docx_to_pdf_fallback_soffice(tmp_path, monkeypatch):
    src = tmp_path / "x.docx"
    _touch(src)
    outdir = tmp_path / "out"
    dst = outdir / "x.pdf"

    monkeypatch.setitem(sys.modules, "docx2pdf", types.SimpleNamespace(convert=lambda *_: (_ for _ in ()).throw(RuntimeError("nope"))))

    # fake soffice
    monkeypatch.setattr(conv.shutil, "which", lambda name: "/usr/bin/soffice")
    def fake_run(cmd, **_):
        produced = Path(cmd[-1]).with_suffix(".pdf")
        (Path(cmd[cmd.index("--outdir")+1]) / produced.name).write_bytes(b"pdf")
        return types.SimpleNamespace(returncode=0, stdout="", stderr="")
    monkeypatch.setattr(conv.subprocess, "run", fake_run)

    res = conv.convert_file(src, dst)
    assert res.engine == "soffice"
    assert dst.exists()

7) pyproject.toml (정적 게이트: ruff/black)
[project]
name = "pdf-word-converter"
version = "0.1.0"
requires-python = ">=3.10"
dependencies = ["pdf2docx>=0.5", "docx2pdf>=0.1", "fastapi>=0.110", "uvicorn>=0.30"]

[tool.black]
line-length = 100

[tool.ruff]
line-length = 100
select = ["E", "F", "I", "UP", "B"]

8) 한계/주의

스캔 PDF(이미지) → DOCX는 텍스트가 빈 채로 이미지만 들어갈 수 있어요. OCR(Tesseract 등) 파이프라인을 추가해야 합니다.

서식 100% 일치는 보장되지 않습니다. 복잡한 표/개체/수식은 어느 쪽이든 후편집이 필요할 수 있어요.

서버/리눅스에서 DOCX→PDF는 Word가 없으므로 LibreOffice 폴백이 정석입니다. soffice --headless --convert-to pdf를 사용합니다.
Ask LibreOffice

9) 2× GitHub 교차검증 (요약)
항목	리포지토리/문서	핵심 관찰	채택/보류
PDF→DOCX	ArtifexSoftware/pdf2docx	PyMuPDF로 PDF 분석 → python-docx로 DOCX 생성. 공식 quickstart 예제와 동일한 API(Converter, parse).
GitHub
+1
	채택: 본 구현에 그대로 사용
DOCX→PDF	AlJohri/docx2pdf	Win/macOS에서 Word 필요, Linux 미지원 명시.
GitHub
+2
PyPI
+2
	조건부 채택: Win/macOS 전용. Linux는 LibreOffice 폴백
폴백 CLI	LibreOffice soffice	--headless --convert-to pdf 표준 호출, 서버 사이드 자동화에 적합.
Ask LibreOffice
	채택
10) 사용 요령 한 장

터미널:

python -m pdf_word_converter.cli ./in/a.pdf ./in/b.docx -o ./out


웹(선택):

uvicorn web.app:app --reload
# POST /convert 로 파일 업로드 → 변환 파일 바로 다운로드


필요하면 OCR 옵션(스캔 PDF 처리)을 붙여서 “PDF 이미지 → 텍스트 있는 DOCX”까지 확장해줄 수도 있어요. 우선은 이 베이스로 굴려보고, 문서 샘플에서 삐끗하는 지점이 보이면 거기부터 다듬죠.
