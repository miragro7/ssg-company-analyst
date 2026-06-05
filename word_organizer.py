#!/usr/bin/env python3
"""
word_organizer.py
임의 파일(docx / txt / html) → sds-word-writer 스타일 Word 파일 변환기

Usage:
    python word_organizer.py <input_file> <output.docx>
"""

import json
import os
import re
import subprocess
import sys
import tempfile
from datetime import datetime, timezone, timedelta
from pathlib import Path

SKILL_DIR = Path(__file__).parent / "skills" / "sds-word-writer" / "scripts"
GENERATE  = str(SKILL_DIR / "generate.py")
KST       = timezone(timedelta(hours=9))
ROMAN     = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ",
             "Ⅺ", "Ⅻ", "ⅩⅢ", "ⅩⅣ", "ⅩⅤ", "ⅩⅥ", "ⅩⅦ", "ⅩⅧ", "ⅩⅨ", "ⅩⅩ"]

# ─────────────────────────────────────────────
# 1. 파일 읽기
# ─────────────────────────────────────────────

def _extract_docx(path: str) -> list:
    """Returns list of (is_heading, text) from docx."""
    from docx import Document
    doc = Document(path)
    result = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style_name = (para.style.name or "").lower()
        is_heading = "heading" in style_name
        if not is_heading and para.runs:
            non_empty = [r for r in para.runs if r.text.strip()]
            if non_empty and all(r.bold for r in non_empty):
                is_heading = True
        result.append((is_heading, text))
    return result


def _extract_html(path: str) -> list:
    """Returns list of (is_heading, text) from html."""
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install",
                               "beautifulsoup4", "-q"])
        from bs4 import BeautifulSoup

    with open(path, encoding="utf-8", errors="replace") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    for tag in soup(["script", "style", "meta", "link"]):
        tag.decompose()

    result = []
    for el in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6",
                              "p", "li", "td", "th"]):
        text = el.get_text(separator=" ", strip=True)
        if not text or len(text) < 2:
            continue
        is_heading = el.name in ("h1", "h2", "h3", "h4", "h5", "h6")
        result.append((is_heading, text))

    # 연속 중복 제거
    deduped, prev = [], None
    for item in result:
        if item[1] != prev:
            deduped.append(item)
            prev = item[1]
    return deduped


def _extract_txt(path: str) -> list:
    """Returns list of (is_heading, text) from plain text."""
    with open(path, encoding="utf-8", errors="replace") as f:
        raw = f.read()

    result = []
    for line in raw.splitlines():
        text = line.strip()
        if not text:
            continue
        is_heading = False
        if len(text) <= 70:
            if re.match(
                r"^(?:\d+\s*[.):\-]\s*|[IVXivx]+\s*[.):\-]\s*|[①-⑳]\s*"
                r"|【|▶|■|◆|▣|\[|\(|제\s*\d+)",
                text
            ):
                is_heading = True
            elif text.endswith(":") and len(text) <= 50:
                is_heading = True
            elif len(text) <= 40 and re.search(r"[A-Z]", text) and text == text.upper():
                is_heading = True
        result.append((is_heading, text))
    return result


# ─────────────────────────────────────────────
# 2. 텍스트 구조 → content.json
# ─────────────────────────────────────────────

def _split_to_bullets(text: str) -> list:
    """긴 텍스트를 불릿 단위로 분리."""
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    if len(lines) > 1:
        return lines
    if len(text) > 200:
        parts = re.split(r"(?<=[.!?。])\s+", text)
        if len(parts) > 1:
            return parts
    return [text]


def _build_content_json(items: list, filename: str) -> dict:
    date_str = datetime.now(KST).strftime("%Y. %m. %d.")

    # 첫 번째 항목을 문서 제목으로 사용
    title = filename
    body_start = 0
    for i, (is_h, text) in enumerate(items):
        if is_h or i == 0:
            title = text
            body_start = i + 1
            break

    # 섹션 그룹핑
    sections = []
    current_heading = None
    current_bullets = []

    def flush():
        if current_heading is not None or current_bullets:
            hdg = current_heading or "내용"
            idx = len(sections)
            numeral = ROMAN[idx] if idx < len(ROMAN) else f"{idx + 1}."
            sections.append({
                "numeral": numeral,
                "heading": hdg,
                "bullets": list(current_bullets),
            })

    for is_h, text in items[body_start:]:
        if is_h:
            flush()
            current_heading = text
            current_bullets = []
        else:
            for sentence in _split_to_bullets(text):
                current_bullets.append({"kind": "dash", "text": sentence})

    flush()

    # 섹션이 없으면 전체를 단일 섹션으로
    if not sections:
        bullets = []
        for _, text in items[body_start:]:
            for s in _split_to_bullets(text):
                bullets.append({"kind": "dash", "text": s})
        if bullets:
            sections.append({
                "numeral": ROMAN[0],
                "heading": "내용",
                "bullets": bullets,
            })

    return {"title": title, "date": date_str, "sections": sections}


# ─────────────────────────────────────────────
# 3. 메인
# ─────────────────────────────────────────────

def organize(input_path: str, output_path: str) -> None:
    ext = Path(input_path).suffix.lower()
    filename = Path(input_path).stem

    if ext == ".docx":
        items = _extract_docx(input_path)
    elif ext in (".html", ".htm"):
        items = _extract_html(input_path)
    elif ext == ".txt":
        items = _extract_txt(input_path)
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext}  (docx / html / txt 만 가능)")

    if not items:
        raise RuntimeError("파일에서 내용을 추출할 수 없습니다.")

    content = _build_content_json(items, filename)

    fd, json_path = tempfile.mkstemp(suffix=".json")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(content, f, ensure_ascii=False, indent=2)

        result = subprocess.run(
            [sys.executable, GENERATE, json_path, output_path],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            raise RuntimeError(f"Word 생성 실패:\n{result.stderr.strip()}")
        print(result.stdout.strip())
    finally:
        try:
            os.unlink(json_path)
        except OSError:
            pass


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python word_organizer.py <input_file> <output.docx>")
        sys.exit(1)
    organize(sys.argv[1], sys.argv[2])
