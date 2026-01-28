---
title: "Development & Tools Ontology"
type: "ontology-design"
domain: "software-development"
sub-domains: ["code-ontology", "document-conversion"]
version: "unified-1.0"
date: "2025-01-19"
tags: ["ontology", "code", "tdd", "development", "tools", "pdf", "docx"]
status: "active"
---

# Part 1: Code Writing Ontology

# CODE 작성 - 온톨로지 관점

## 요약

"코드 작성"을 **온톨로지**(무엇이 있고 어떻게 얽히는가)로 정의합니다.

이 프로젝트는 **원칙(Constitution)**이 **작업 흐름(Commands)**을 **산출물(Artifacts)**로 이어 붙이고, 각 흐름은 **게이트(Quality Gates)**로 검증되며, 모든 조각이 **명세–계획–과업–분석–구현**이라는 **유형화된 과정 타입들**로 상호 제약되는 체계입니다. 헌법 규칙은 상위 규범이며 분석 단계에서 "불가침"으로 집행됩니다.

__1\) TBox\(스키마\): 핵심 클래스와 관계__

__클래스\(개념\)__

- __Project__: 코드 작성 전체를 아우르는 최상위 개체\.
- __Principle__: 변경 불가한 규범\. *Constitution* 문서로 구체화된다\(예: 의사결정·품질 기준\)\.
- __Command\(ProcessType\)__: /specify·/clarify·/plan·/tasks·/analyze·/implement 같은 실행 가능한 과정 유형\. 각 커맨드는 입력/출력 산출물이 정형화되어 있다\.
- __Artifact__: spec\.md, plan\.md, tasks\.md, research\.md, data\-model\.md, contracts/, quickstart\.md 같은 파일 산출물\.
- __QualityGate__: “진행해도 되는가?”를 판단하는 규범/검증 모음\. 예: 헌법 준수 체크, TDD/Tidy 루프, 선행 명세/정리 유무\.
- __State__: Artifact/Feature의 생애주기 상태\(초안 Draft → Clarified → Planned → Tasked → Analyzed → Implemented\)\.
- __Actor__: 참여 주체\(개발자/리뷰어/자동 에이전트\)\.
- __Tool__: 테스트/린트·스크립트 등\(예: 계획/검증 스크립트, 테스트 러너\)\. 커맨드가 이를 호출한다\.

__핵심 관계__

- governs\(Principle, Command/Artifact\): 헌법 원칙이 프로세스와 산출물을 규율\.
- produces\(Command, Artifact\): 각 커맨드가 특정 산출물을 생성/갱신\.
- requires\(Command, Artifact|Principle\): 실행 전제조건\(예: /plan은 __Clarifications__가 있어야 함\)\.
- validates\(QualityGate, Artifact|Process\): 분석/게이트가 요구사항 충족 여부를 판정\(헌법 충돌은 자동 *CRITICAL*\)\.
- orders\(ProcessType, ProcessType\): 수행 순서 제약\(예: 테스트→코드, 구조→행위\)\.

__2\) 프로세스 타입\(Commands\)로 본 수명주기__

1. __/specify__ → *특징*: 브랜치와 스펙 파일을 생성·체크아웃하고 템플릿에 맞춰 __spec\.md__를 초기화한다\. 산출: SPEC\_FILE, 분기: BRANCH\_NAME\.
2. __/clarify__ → *역할*: 최대 5개의 정밀 질문으로 __모호성 해소__ 후, 답을 __Clarifications__ 섹션과 관련 섹션\(요구사항·데이터모델·품질 속성 등\)에 __즉시 반영__한다\. 상태 전이: Draft → Clarified\.
3. __/plan__ → *전제*: Clarifications가 존재해야 진행\. *산출*: __research\.md__, __data\-model\.md__, __contracts/__, __quickstart\.md__, 그리고 최종적으로 __tasks\.md__\(Phase2\)\.
4. __/tasks__ → *규칙*: 의존성 순서와 병렬성\(\[P\]\)을 명시\. “__테스트 우선\(TDD\)__, 모델→서비스→엔드포인트, 코어→통합→폴리시” 등 정형 규칙을 반영해 __실행가능한__ 작업 목록을 생성\.
5. __/analyze__ → *성격*: __읽기 전용__ 교차검증\. Spec/Plan/Tasks 간 __중복·모호·미커버리지·헌법 충돌__을 표 형태로 보고하고 심각도\(CRITICAL/HIGH/…\)를 부여\. 헌법 위반은 자동 CRITICAL\.
6. __/implement__ → *집행*: __tasks\.md__의 단계/의존/병렬 규칙에 따라 __단계별__ 구현\. 실패 시\(비병렬\) 중단·보고, 병렬은 성공분 계속\. 완료 시 체크박스를 \[X\]로 갱신\.

__3\) 규범\(Principles\)과 게이트\(Quality Gates\)__

- __TDD 루프\(RED→GREEN→REFACTOR\)__: 기능 한 조각의 실패 테스트\(RED\) → 통과에 필요한 __최소 구현__\(GREEN\) → __중복 제거/의도 드러내기__\(REFACTOR\)\. 단일 테스트는 ms~수백 ms SLA, 느린 I/O는 더블/주입\.
- __Tidy First\(구조→행위 분리\)__: 리네임/추출/이동 등 __구조 변경__은 행위 불변으로 먼저, 그 다음 __행위 변경__\(feat/fix\)\. 커밋도 structural: ↔ behavioral:로 분리\.
- __진입 게이트__: /plan은 명시적으로 __Clarifications__ 유무를 검사\(없으면 중단·사전 정리 권고\)\. /analyze는 헌법 위반을 __CRITICAL__로 지정\.

이 세 가지는 온톨로지 상에서 QualityGate 인스턴스들로 모델링되어 각 Process 인스턴스에 연결된다\. 실패 시 __State__는 다음 단계로 전이되지 않는다\(불합격 전이 차단\)\.

__4\) ABox\(예시 인스턴스\) — 한 기능이 지나가는 이야기__

- Process\(/specify\) __produces__ Artifact\(spec\.md\) & __creates__ Branch\(feature/x\)\.
- Process\(/clarify\) __updates__ Artifact\(spec\.md\#Clarifications\); 상태 Draft→Clarified\.
- Process\(/plan\) __produces__ research\.md, data\-model\.md, contracts/, quickstart\.md, tasks\.md\.
- Process\(/tasks\) __orders__ 작업: __Setup→Tests→Core→Integration→Polish__, 파일 충돌 시 __순차__, 파일 분리 시 __\[P\] 병렬__\.
- Process\(/analyze\) __validates__ \(dup/ambiguity/coverage/constitution\) ⇒ __OK__면 다음\.
- Process\(/implement\) __executes__ tasks; 실패 시 중단·리포트, 성공은 \[X\] 체크\.

__5\) 소형 온톨로지 스케치\(의미 보존용 의사‑Turtle\)__

\#\#\# Classes

Project, Principle, Command, Artifact, QualityGate, State \.

\#\#\# Object properties

governs, produces, requires, validates, orders, updates, creates, blocksTransition \.

\#\#\# Instances \(축약 예\)

:Constitution        a Principle \.

:TDD\_Gate            a QualityGate \.

:TidyFirst\_Gate      a QualityGate \.

:CodeWriting         a Project \.

:Specify  a Command ; produces :SpecMD ; creates :FeatureBranch \.

:Clarify  a Command ; requires :SpecMD ; updates :SpecMD \.

:Plan     a Command ; requires :SpecMD ; produces :Research, :DataModel, :Contracts, :Quickstart, :TasksMD \.

:Tasks    a Command ; requires :Plan ; produces :TasksMD \.

:Analyze  a Command ; validates :SpecMD, :PlanMD, :TasksMD ; governedBy :Constitution \.

:Implement a Command ; requires :TasksMD ; validates :TDD\_Gate, :TidyFirst\_Gate \.

\# 규칙 예시

:Analyze \-\-governedBy\-\-> :Constitution \.   \# 헌법 불가침

:TDD\_Gate \-\-blocksTransition\-\-> :ImplementWhenFailed \.

- 헌법 규칙의 “불가침”과 분석의 읽기 전용·심각도 체계는 /analyze 문서의 규정으로 근거화된다\.
- Clarify의 “최대 5문항”·분류 택소노미·즉시 반영 규칙은 /clarify 흐름 정의로부터 온다\.
- Plan의 단계 산출물 세트와 Tasks의 의존·병렬 규칙은 각 명세에 기재되어 있다\.
- TDD/Tidy 루프와 구조↔행위 분리는 Tidy 지침에 명시되어 있다\.

__6\) 이 온톨로지가 답해야 하는 Competency Questions__

- 현재 Feature X는 __어떤 상태__이며\(Clarified? Planned?\) 다음 단계로 __전이 가능한가?__ 전이를 막는 __게이트 실패__는 무엇인가? \(예: Clarifications 없음, Analyze CRITICAL\)
- spec/plan/tasks 사이에 __중복·충돌·미커버리지__가 있는가? 있다면 __심각도__와 __수정 후보__는?
- 특정 과업 T\#\#\#는 __어떤 요구사항__을 충족하며, __병렬 실행__ 가능한가\(파일 경합 여부\)?
- 이번 커밋은 __구조 변경__인가 __행위 변경__인가? \(분리 커밋/리뷰 기준\)

__7\) 함께 볼 메타‑원칙\(현실과의 접점\)__

- __plan\.md 최우선 & RED→GREEN→REFACTOR & 구조→행위__: 설계→검증→집행 순서가 지식 구조로서 먼저 서야, 실제 코드 변화가 안정적으로 수렴한다\.
- __/plan 전 Clarifications 필수__: 모호성은 가장 비싼 결함\. 전방에서 줄인다\.
- __헌법 우선권__: 원칙을 바꾸려면 별도 절차\(버전/개정 내역/동기\)가 필요하다\.

__다음 선택지\(가벼운 옵션\)__

1. 위 스키마를 __TTL/OWL 초안__으로 뽑아드릴 수 있다\(프로세스·산출물·게이트·상태 포함\)\.
2. 현재 진행 중인 기능 하나를 골라 __ABox 인스턴스__\(실제 삼중\)로 채워 시각화해보자\.
3. /analyze 리포트 항목을 __온톨로지 질의__\(competency questions\)로 바로 매핑하는 규칙표를 만들어 드릴 수도 있다\.

원하면 바로 골라줘\. 난 틀을 잡아주고, 너는 도메인 언어를 더해 주면 된다\.
이 구조는 네가 원하는 “온톨로지 관점의 설계—운영—검증”을 같은 축 위에 올려준다\.

---

# Part 2: PDF-DOCX Bidirectional Converter

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
