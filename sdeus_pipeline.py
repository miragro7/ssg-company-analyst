#!/usr/bin/env python3
"""
스드스(Sdeus) 통합 파이프라인
기업 데이터 딕셔너리 → HTML 리포트 + SDS Word 문서 생성 → 텔레그램 2파일 전송
"""

import json
import os
import subprocess
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path

# Load .env if present
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

BASE_DIR   = Path(__file__).resolve().parent
SKILL_DIR  = BASE_DIR / "skills" / "sds-word-writer" / "scripts"
REPORT_DIR = Path(os.environ.get("REPORT_DIR", str(BASE_DIR / "report")))
BOT_TOKEN  = os.environ.get("BOT_TOKEN", "")
CHAT_ID    = int(os.environ.get("CHAT_ID", "0"))

KST  = timezone(timedelta(hours=9))
ROMAN = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ", "Ⅺ", "Ⅻ", "ⅩⅢ", "ⅩⅣ"]

sys.path.insert(0, str(BASE_DIR))
from sdeus_reporter import save_report

import telebot


# ───────────────────────────────────────────────────────────────
# 헬퍼
# ───────────────────────────────────────────────────────────────

def _bullets(text: str, kind: str = "dot") -> list:
    return [{"kind": kind, "text": line.strip()}
            for line in (text or "").split("\n") if line.strip()]


def _dash(text: str) -> dict:
    return {"kind": "dash", "text": text}


def _dot(text: str) -> dict:
    return {"kind": "dot", "text": text}


def _empty() -> list:
    return [{"kind": "dash", "text": "데이터 없음"}]


# ───────────────────────────────────────────────────────────────
# 1. data dict → content.json 변환 (Word 생성용)
# ───────────────────────────────────────────────────────────────

def data_to_content_json(data: dict) -> dict:
    company    = data.get("company_name", "기업명")
    company_en = data.get("company_name_en", "")
    ticker     = data.get("ticker", "-")
    exchange   = data.get("exchange", "-")
    concept    = data.get("concept", "종합분석")
    report_date = data.get("report_date", datetime.now(KST).strftime("%Y. %m. %d."))
    date_fmt   = report_date.replace("-", ". ")
    if not date_fmt.endswith("."):
        date_fmt += "."

    ov   = data.get("overview", {})
    mk   = data.get("market", {})
    drv  = data.get("growth_drivers", [])
    fin  = data.get("financials", [])
    ofc  = data.get("offices", [])
    exc  = data.get("executives", [])
    refs = data.get("references", [])

    es   = data.get("executive_summary", {})
    comp = data.get("competitive", {})
    bm   = data.get("business_model", {})
    finx = data.get("financial_extended", {})
    risk = data.get("risk", {})
    val  = data.get("valuation", {})
    its  = data.get("investment_thesis", [])
    scen = data.get("scenario", {})
    conc = data.get("conclusion", {})

    toc_items = [
        "Executive Summary (핵심 요약)",
        "Company Overview (기업 개요)",
        "Industry Analysis (산업 분석)",
        "Competitive Landscape (경쟁 구도)",
        "Business Model Analysis (사업모델)",
        "Management Analysis (경영진)",
        "Financial Analysis (재무 분석)",
        "Growth Drivers (성장동력)",
        "Risk Analysis (리스크 분석)",
        "Valuation (기업가치 평가)",
        "Investment Thesis (투자포인트)",
        "Scenario Analysis (시나리오)",
        "Conclusion (결론)",
        "자료 출처",
    ]

    sections = []

    # ── Ⅰ. Executive Summary ──
    if es:
        s_bullets = [
            _dash(f"기업 한 줄 정의 : {es.get('one_line', '-')}"),
            _dash(f"핵심 경쟁력 : {es.get('key_competency', '-')}"),
            _dash(f"현재 위치 : {es.get('current_position', '-')}"),
            _dash(f"향후 전망 : {es.get('outlook', '-')}"),
            _dash("결론"),
        ] + _bullets(es.get("conclusion", ""), "dot")
    else:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[0], "heading": "Executive Summary (핵심 요약)", "bullets": s_bullets})

    # ── Ⅱ. Company Overview ──
    co_name = f"{company}{(' (' + company_en + ')') if company_en else ''}"
    s_bullets = [
        _dash(f"기업명 : {co_name}"),
        _dash(f"티커 / 거래소 : {ticker} ({exchange})"),
        _dash(f"설립연도 : {ov.get('founded', '-')}"),
        _dash(f"본사 위치 : {ov.get('hq', '-')}"),
        _dash(f"주요 산업 : {ov.get('industry', '-')}"),
        _dash(f"임직원 수 : {ov.get('employees', '-')}"),
        _dash("기업 개요"),
    ] + _bullets(ov.get("summary", ""), "dot")
    sections.append({"numeral": ROMAN[1], "heading": "Company Overview (기업 개요)", "bullets": s_bullets})

    # ── Ⅲ. Industry Analysis ──
    s_bullets = [
        _dash(f"시장 규모 (TAM) : {mk.get('total_size', '-')}"),
        _dash(f"연평균 성장률 : {mk.get('growth_rate', '-')}"),
        _dash(f"시장 점유율 : {mk.get('market_share', '-')}"),
        _dash("산업 트렌드"),
    ] + _bullets(mk.get("trend", ""), "dot")
    sections.append({"numeral": ROMAN[2], "heading": "Industry Analysis (산업 분석)", "bullets": s_bullets})

    # ── Ⅳ. Competitive Landscape ──
    if comp:
        comp_rows = [[c.get("name","-"), c.get("revenue","-"),
                      c.get("market_share","-"), c.get("strategy","-")]
                     for c in comp.get("competitors", [])]
        s_bullets = [
            {"kind": "table",
             "caption": "주요 경쟁사 비교",
             "columns": ["기업명", "매출", "시장 점유율", "핵심 전략"],
             "rows": comp_rows} if comp_rows else _dash("경쟁사 데이터 없음"),
            _dash("포지셔닝 분석"),
        ] + _bullets(comp.get("positioning", ""), "dot") + [
            _dash(f"강점 : {comp.get('strengths', '-')}"),
            _dash(f"약점 : {comp.get('weaknesses', '-')}"),
        ]
    else:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[3], "heading": "Competitive Landscape (경쟁 구도)", "bullets": s_bullets})

    # ── Ⅴ. Business Model Analysis ──
    if bm:
        s_bullets = [
            _dash("수익 구조"),
        ] + _bullets(bm.get("revenue_structure", ""), "dot") + [
            _dash("비용 구조"),
        ] + _bullets(bm.get("cost_structure", ""), "dot") + [
            _dash("수익성 구조"),
        ] + _bullets(bm.get("profitability", ""), "dot") + [
            _dash("경쟁 우위"),
        ] + _bullets(bm.get("competitive_advantage", ""), "dot")
    else:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[4], "heading": "Business Model Analysis (사업모델)", "bullets": s_bullets})

    # ── Ⅵ. Management Analysis ──
    s_bullets = []
    for e in exc:
        s_bullets.append(_dash(f"{e.get('name', '-')} — {e.get('title', '-')}"))
        s_bullets.append(_dot(f"학력 : {e.get('education', '-')}"))
        for line in e.get("career", "").split("\n"):
            line = line.strip()
            if line:
                s_bullets.append(_dot(line if line.startswith("이력") else f"이력 : {line}"))
    if not s_bullets:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[5], "heading": "Management Analysis (경영진 분석)", "bullets": s_bullets})

    # ── Ⅶ. Financial Analysis ──
    fin_rows = [[f.get("year","-"), f.get("revenue","-"), f.get("op_income","-"),
                 f.get("net_income","-"), f.get("eps","-"), f.get("roe","-")] for f in fin]
    bs  = finx.get("balance_sheet", {})
    cf  = finx.get("cashflow", {})
    kpi = finx.get("kpi", {})
    s_bullets = [
        _dash("손익계산서 (5개년)"),
        {"kind": "table",
         "caption": f"{company} 5개년 재무 실적",
         "columns": ["연도", "매출액", "영업이익", "순이익", "EPS", "ROE"],
         "rows": fin_rows},
    ]
    if bs:
        s_bullets += [
            _dash("재무상태표"),
            _dot(f"총자산 : {bs.get('assets', '-')}"),
            _dot(f"총부채 : {bs.get('liabilities', '-')}"),
            _dot(f"자기자본 : {bs.get('equity', '-')}"),
        ]
    if cf:
        s_bullets += [
            _dash("현금흐름"),
            _dot(f"영업현금흐름 : {cf.get('operating', '-')}"),
            _dot(f"투자현금흐름 : {cf.get('investing', '-')}"),
            _dot(f"잉여현금흐름(FCF) : {cf.get('fcf', '-')}"),
        ]
    if kpi:
        s_bullets.append(_dash("주요 지표"))
        for lbl, key in [("ROE", "roe"), ("ROIC", "roic"), ("EBITDA", "ebitda"), ("영업이익률", "op_margin"), ("부채비율", "debt_ratio")]:
            if kpi.get(key):
                s_bullets.append(_dot(f"{lbl} : {kpi[key]}"))
    sections.append({"numeral": ROMAN[6], "heading": "Financial Analysis (재무 분석)", "bullets": s_bullets})

    # ── Ⅷ. Growth Drivers ──
    s_bullets = []
    for i, d in enumerate(drv):
        s_bullets.append(_dash(f"{'①②③④⑤⑥⑦⑧⑨'[i] if i < 9 else str(i+1)} {d.get('title', '')}"))
        s_bullets += _bullets(d.get("detail", ""), "dot")
    if not s_bullets:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[7], "heading": "Growth Drivers (성장동력)", "bullets": s_bullets})

    # ── Ⅸ. Risk Analysis ──
    if risk:
        s_bullets = []
        for lbl, key in [("사업 리스크", "business"), ("기술 리스크", "technology"),
                         ("규제 리스크", "regulatory"), ("경쟁 리스크", "competitive"),
                         ("거시경제 리스크", "macro")]:
            if risk.get(key):
                s_bullets.append(_dash(lbl))
                s_bullets += _bullets(risk[key], "dot")
    else:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[8], "heading": "Risk Analysis (리스크 분석)", "bullets": s_bullets})

    # ── Ⅹ. Valuation ──
    if val:
        is_listed = val.get("is_listed", True)
        rel_v = val.get("relative", {})
        abs_v = val.get("absolute", {})
        s_bullets = [_dash("상장 기업" if is_listed else "비상장 기업 — 상대배수 일부 미적용")]
        if is_listed:
            for lbl, key in [("PER", "per"), ("PBR", "pbr"), ("EV/EBITDA", "ev_ebitda")]:
                if rel_v.get(key):
                    s_bullets.append(_dot(f"{lbl} : {rel_v[key]}"))
        if abs_v.get("dcf"):
            s_bullets.append(_dot(f"DCF : {abs_v['dcf']}"))
        if val.get("target_value"):
            s_bullets.append(_dot(f"목표 기업가치 : {val['target_value']}"))
        if is_listed and val.get("target_price"):
            s_bullets.append(_dot(f"적정 주가 : {val['target_price']}"))
    else:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[9], "heading": "Valuation (기업가치 평가)", "bullets": s_bullets})

    # ── Ⅺ. Investment Thesis ──
    if its:
        s_bullets = []
        for i, t in enumerate(its):
            s_bullets.append(_dash(f"{'①②③④⑤'[i] if i < 5 else str(i+1)} {t.get('point', '')}"))
            s_bullets += _bullets(t.get("detail", ""), "dot")
    else:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[10], "heading": "Investment Thesis (투자포인트)", "bullets": s_bullets})

    # ── Ⅻ. Scenario Analysis ──
    if scen:
        s_bullets = []
        for label, key in [("Bull Case (낙관)", "bull"), ("Base Case (기준)", "base"), ("Bear Case (비관)", "bear")]:
            s = scen.get(key, {})
            if s:
                s_bullets.append(_dash(label))
                if s.get("conditions"):
                    s_bullets.append(_dot(f"조건 : {s['conditions']}"))
                if s.get("outlook"):
                    s_bullets.append(_dot(f"전망 : {s['outlook']}"))
    else:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[11], "heading": "Scenario Analysis (시나리오 분석)", "bullets": s_bullets})

    # ── ⅩⅢ. Conclusion ──
    if conc:
        s_bullets = [
            _dash("최종 평가"),
        ] + _bullets(conc.get("final_assessment", ""), "dot") + [
            _dash("핵심 체크포인트"),
        ] + _bullets(conc.get("checkpoints", ""), "dot") + [
            _dash("향후 관전 포인트"),
        ] + _bullets(conc.get("watchpoints", ""), "dot")
    else:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[12], "heading": "Conclusion (결론)", "bullets": s_bullets})

    # ── ⅩⅣ. 자료 출처 ──
    s_bullets = []
    for r in refs:
        s_bullets.append(_dash(r.get("title", r.get("url", ""))))
        s_bullets.append(_dot(f"URL : {r.get('url', '-')}"))
    if not s_bullets:
        s_bullets = _empty()
    sections.append({"numeral": ROMAN[13], "heading": "자료 출처", "bullets": s_bullets})

    return {
        "title": f"({concept}) {company} 기업 분석 보고서",
        "date": date_fmt,
        "toc": {
            "reportTitle": f"({concept}) {company} 기업 분석 보고서",
            "items": toc_items,
            "date": date_fmt,
            "author": "스드스(Sdeus) 기업 분석 시스템",
        },
        "sections": sections,
    }


# ───────────────────────────────────────────────────────────────
# 2. Word 생성
# ───────────────────────────────────────────────────────────────

def save_word(data: dict, html_path: str) -> str:
    base = os.path.splitext(html_path)[0]
    docx_path         = base + ".docx"
    content_json_path = base + "_content.json"

    content = data_to_content_json(data)
    with open(content_json_path, "w", encoding="utf-8") as f:
        json.dump(content, f, ensure_ascii=False, indent=2)

    result = subprocess.run(
        [sys.executable, str(SKILL_DIR / "generate.py"), content_json_path, docx_path],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"Word 생성 실패: {result.stderr}")
    print(f"[스드스] Word 저장 완료: {docx_path}")
    return docx_path


# ───────────────────────────────────────────────────────────────
# 3. 텔레그램 전송 (HTML + DOCX 2파일)
# ───────────────────────────────────────────────────────────────

def send_both(company: str, html_path: str, docx_path: str):
    if not BOT_TOKEN or not CHAT_ID:
        print("[스드스] BOT_TOKEN 또는 CHAT_ID 미설정 — 텔레그램 전송 건너뜀")
        return

    bot = telebot.TeleBot(BOT_TOKEN)

    with open(html_path, "rb") as f:
        bot.send_document(
            CHAT_ID, f,
            caption=f"📊 [{company}] 기업 분석 리포트 (HTML) — 스드스(Sdeus)",
            visible_file_name=os.path.basename(html_path),
        )
    print(f"[스드스] HTML 전송 완료: {html_path}")

    with open(docx_path, "rb") as f:
        bot.send_document(
            CHAT_ID, f,
            caption=f"📄 [{company}] 기업 분석 보고서 (Word) — 스드스(Sdeus)",
            visible_file_name=os.path.basename(docx_path),
        )
    print(f"[스드스] Word 전송 완료: {docx_path}")


# ───────────────────────────────────────────────────────────────
# 4. 통합 실행 함수 (외부 import 용)
# ───────────────────────────────────────────────────────────────

def run_pipeline(data: dict):
    """
    기업 분석 전체 파이프라인 실행.
    1) HTML 리포트 생성
    2) SDS Word 문서 생성
    3) 텔레그램으로 두 파일 전송
    """
    company = data.get("company_name", "unknown")
    print(f"\n[스드스] ── {company} 분석 파이프라인 시작 ──")
    html_path = save_report(data)
    docx_path = save_word(data, html_path)
    send_both(company, html_path, docx_path)
    print(f"[스드스] ── {company} 분석 파이프라인 완료 ──\n")
    return {"html": html_path, "docx": docx_path}


# ───────────────────────────────────────────────────────────────
# 직접 실행 시 : 최근 리포트 재전송 테스트
# ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import glob
    html_files = sorted(glob.glob(str(REPORT_DIR / "*.html")))
    if not html_files:
        print("테스트할 HTML 파일 없음")
        sys.exit(1)

    html_path = html_files[-1]
    docx_path = html_path.replace(".html", ".docx")
    company   = Path(html_path).stem.rsplit("_", 2)[0]

    if not os.path.exists(docx_path):
        print(f"Word 파일 없음: {docx_path}")
        sys.exit(1)

    print(f"[테스트] HTML: {html_path}")
    print(f"[테스트] DOCX: {docx_path}")
    send_both(company, html_path, docx_path)
