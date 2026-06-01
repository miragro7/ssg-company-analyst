#!/usr/bin/env python3
"""스드스 기업 분석 HTML 리포트 생성기 — 14섹션 템플릿"""

import os
from datetime import datetime, timezone, timedelta
from pathlib import Path

KST = timezone(timedelta(hours=9))
REPORT_DIR = os.environ.get("REPORT_DIR", str(Path(__file__).resolve().parent / "report"))

CONCEPT_META = {
    "투자판단":       {"label": "투자판단 중심",       "color": "#0d6efd"},
    "재무건전성":     {"label": "재무건전성 중심",     "color": "#198754"},
    "성장전략":       {"label": "성장전략 중심",       "color": "#6f42c1"},
    "사업성과리스크": {"label": "사업성과·리스크 중심", "color": "#dc3545"},
    "종합분석":       {"label": "종합 분석",           "color": "#1a3a5c"},
}


def kst_now():
    return datetime.now(KST)


def _nl(text: str) -> str:
    return (text or "").replace("\n", "<br>")


def _empty_section(sec_id: str, title: str) -> str:
    return f'<section id="{sec_id}" class="section"><h2 class="section-title">{title}</h2><p class="empty">데이터 없음</p></section>'


def _office_type(t: str) -> str:
    if "본사" in t: return "hq"
    if any(x in t for x in ("지사", "사무소", "센터")): return "branch"
    if any(x in t for x in ("공장", "생산", "플랜트")): return "factory"
    return "other"


def build_html_report(data: dict) -> str:
    company    = data.get("company_name", "기업명 미입력")
    company_en = data.get("company_name_en", "")
    ticker     = data.get("ticker", "-")
    exchange   = data.get("exchange", "-")
    report_date = data.get("report_date", kst_now().strftime("%Y-%m-%d"))
    concept    = data.get("concept", "종합분석")
    cm         = CONCEPT_META.get(concept, CONCEPT_META["종합분석"])

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

    co   = f"{company}{(' (' + company_en + ')') if company_en else ''}"
    title = f"({cm['label']}) {company} 기업 분석 보고서"
    c = cm["color"]

    # ── CSS ──────────────────────────────────────────────────────
    css = f"""
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', Arial, sans-serif;
      background: #f0f2f5; color: #1e2328;
      padding: 24px 16px; font-size: 15px; line-height: 1.75;
    }}
    a {{ color: #1a6fcf; text-decoration: none; }}
    a:hover {{ text-decoration: underline; }}
    .container {{ max-width: 960px; margin: 0 auto; }}
    .report-header {{
      background: linear-gradient(135deg, #1a3a5c 0%, {c} 100%);
      color: #fff; border-radius: 12px; padding: 32px 36px; margin-bottom: 28px;
    }}
    .badge-concept {{
      display: inline-block; background: rgba(255,255,255,0.2); color: #fff;
      font-size: 0.78em; padding: 3px 12px; border-radius: 20px;
      margin-bottom: 10px; letter-spacing: 0.05em;
    }}
    .report-header h1 {{ font-size: 1.5em; font-weight: 700; line-height: 1.4; margin-bottom: 8px; }}
    .report-header .meta {{ font-size: 0.85em; opacity: 0.85; }}
    .toc {{
      background: #fff; border-radius: 10px; padding: 20px 28px;
      margin-bottom: 24px; box-shadow: 0 1px 4px rgba(0,0,0,0.07);
    }}
    .toc-title {{ font-size: 1em; font-weight: 700; color: #1a3a5c; margin-bottom: 12px; }}
    .toc-list {{ columns: 2; padding-left: 20px; color: #444; font-size: 0.9em; line-height: 2.1; }}
    .toc-list a {{ color: #1a6fcf; }}
    .section {{
      background: #fff; border-radius: 10px; padding: 28px 32px;
      margin-bottom: 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.07);
    }}
    .section-title {{
      font-size: 1.05em; font-weight: 700; color: #1a3a5c;
      border-left: 4px solid {c}; padding-left: 12px; margin-bottom: 18px;
    }}
    .sub-title {{
      font-size: 0.9em; font-weight: 700; color: #1a3a5c;
      margin: 18px 0 10px; border-bottom: 1px dashed #ccd; padding-bottom: 6px;
    }}
    .info-table {{ width: 100%; border-collapse: collapse; font-size: 0.9em; margin-bottom: 14px; }}
    .info-table th {{
      background: #f0f4fa; color: #1a3a5c; font-weight: 600;
      text-align: left; padding: 10px 14px; width: 28%; border: 1px solid #d8e0ed;
    }}
    .info-table td {{ padding: 10px 14px; border: 1px solid #d8e0ed; color: #333; }}
    .text-box {{
      background: #f8f9fb; border-left: 4px solid {c};
      border-radius: 0 6px 6px 0; padding: 14px 18px;
      font-size: 0.9em; color: #444; line-height: 1.9; margin-top: 8px;
    }}
    .highlight-box {{ border-left-color: #e65c00; background: #fff8f4; }}
    .two-col {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }}
    .three-col {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 14px; }}
    .card {{
      background: #f8f9fb; border-radius: 8px; padding: 16px 18px;
      border-top: 3px solid {c};
    }}
    .card-label {{ font-size: 0.78em; font-weight: 700; color: {c}; margin-bottom: 6px; letter-spacing: 0.04em; }}
    .card-value {{ font-size: 0.9em; color: #222; line-height: 1.85; }}
    .card-left {{ border-top: none; border-left: 3px solid {c}; }}
    .es-conclusion {{
      background: {c}; color: #fff; border-radius: 8px;
      padding: 18px 22px; margin-top: 14px; font-size: 0.95em; line-height: 1.8;
    }}
    .table-wrap {{ overflow-x: auto; }}
    .data-table {{ width: 100%; border-collapse: collapse; font-size: 0.88em; }}
    .data-table thead tr {{ background: #1a3a5c; color: #fff; }}
    .data-table th {{ padding: 10px 14px; text-align: center; font-weight: 600; }}
    .data-table td {{ padding: 9px 14px; text-align: center; border-bottom: 1px solid #e0e6ee; color: #333; }}
    .data-table td.left {{ text-align: left; }}
    .data-table tbody tr:nth-child(even) {{ background: #f5f7fb; }}
    .data-table tbody tr:hover {{ background: #eaf0fb; }}
    .data-table .highlight-row {{ background: #e8f0fe !important; font-weight: 700; }}
    .year-cell {{ font-weight: 700; color: #1a3a5c; }}
    .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 12px; margin-top: 14px; }}
    .kpi-card {{
      background: #f0f4fa; border-radius: 8px; padding: 14px 16px; text-align: center;
      border-bottom: 3px solid {c};
    }}
    .kpi-label {{ font-size: 0.78em; color: #666; font-weight: 600; margin-bottom: 6px; }}
    .kpi-value {{ font-size: 1.1em; font-weight: 700; color: #1a3a5c; }}
    .badge {{ display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 0.82em; font-weight: 600; white-space: nowrap; }}
    .badge-hq      {{ background: #dbeafe; color: #1e40af; }}
    .badge-branch  {{ background: #dcfce7; color: #15803d; }}
    .badge-factory {{ background: #fef3c7; color: #92400e; }}
    .badge-other   {{ background: #f3e8ff; color: #6b21a8; }}
    .exec-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 16px; }}
    .exec-card {{ background: #f8f9fb; border-radius: 8px; padding: 18px 20px; border-top: 3px solid {c}; }}
    .exec-name {{ font-size: 1.05em; font-weight: 700; color: #1a3a5c; margin-bottom: 4px; }}
    .exec-title-tag {{ font-size: 0.85em; color: {c}; font-weight: 600; margin-bottom: 10px; }}
    .exec-detail {{ font-size: 0.85em; color: #555; line-height: 1.9; }}
    .exec-label {{
      display: inline-block; background: #e0e8f5; color: #1a3a5c;
      font-weight: 600; font-size: 0.82em; padding: 1px 7px; border-radius: 4px; margin-right: 4px;
    }}
    .driver-card {{
      background: #f8f9fb; border-radius: 8px; padding: 16px 20px;
      margin-bottom: 12px; border-left: 3px solid {c};
    }}
    .driver-title {{ font-weight: 700; color: #1a3a5c; margin-bottom: 6px; font-size: 0.95em; }}
    .driver-detail {{ font-size: 0.9em; color: #555; line-height: 1.8; }}
    .risk-card {{ background: #fff5f5; border-radius: 8px; padding: 16px 18px; border-left: 3px solid #dc3545; }}
    .risk-card-title {{ font-size: 0.85em; font-weight: 700; color: #dc3545; margin-bottom: 8px; }}
    .risk-card-body {{ font-size: 0.88em; color: #555; line-height: 1.8; }}
    .listed-tag {{
      display: inline-block; font-size: 0.82em; color: #666;
      border: 1px solid #ddd; border-radius: 6px; padding: 3px 10px; margin-bottom: 12px;
    }}
    .thesis-list {{ list-style: none; padding: 0; }}
    .thesis-item {{
      display: flex; align-items: flex-start; gap: 14px;
      background: #f8f9fb; border-radius: 8px; padding: 16px 18px; margin-bottom: 12px;
    }}
    .thesis-num {{
      flex-shrink: 0; width: 32px; height: 32px; border-radius: 50%;
      background: {c}; color: #fff; font-weight: 700; font-size: 0.9em;
      display: flex; align-items: center; justify-content: center;
    }}
    .thesis-text {{ flex: 1; }}
    .thesis-point {{ font-weight: 700; color: #1a3a5c; margin-bottom: 4px; font-size: 0.95em; }}
    .thesis-detail {{ font-size: 0.88em; color: #555; line-height: 1.8; }}
    .scenario-card {{ border-radius: 8px; padding: 18px; }}
    .scenario-bull {{ background: #f0fdf4; border-top: 4px solid #16a34a; }}
    .scenario-base {{ background: #eff6ff; border-top: 4px solid #2563eb; }}
    .scenario-bear {{ background: #fff5f5; border-top: 4px solid #dc2626; }}
    .scenario-label {{ font-size: 0.88em; font-weight: 700; margin-bottom: 10px; }}
    .label-bull {{ color: #16a34a; }}
    .label-base {{ color: #2563eb; }}
    .label-bear {{ color: #dc2626; }}
    .scenario-body {{ font-size: 0.88em; color: #444; line-height: 1.85; }}
    .ref-list {{ list-style: none; padding: 0; }}
    .ref-list li {{ padding: 8px 0; border-bottom: 1px solid #eee; font-size: 0.88em; }}
    .ref-list li:last-child {{ border-bottom: none; }}
    .ref-list a {{ word-break: break-all; }}
    .report-footer {{ text-align: center; color: #888; font-size: 0.82em; padding: 24px 0 8px; }}
    .empty {{ color: #aaa; font-style: italic; font-size: 0.88em; }}
    """

    # ── 목차 ──────────────────────────────────────────────────────
    toc_html = """
    <nav class="toc">
      <h2 class="toc-title">📋 목차</h2>
      <ol class="toc-list">
        <li><a href="#sec-es">Executive Summary (핵심 요약)</a></li>
        <li><a href="#sec-overview">Company Overview (기업 개요)</a></li>
        <li><a href="#sec-industry">Industry Analysis (산업 분석)</a></li>
        <li><a href="#sec-competitive">Competitive Landscape (경쟁 구도)</a></li>
        <li><a href="#sec-bm">Business Model Analysis (사업모델)</a></li>
        <li><a href="#sec-mgmt">Management Analysis (경영진)</a></li>
        <li><a href="#sec-financial">Financial Analysis (재무 분석)</a></li>
        <li><a href="#sec-growth">Growth Drivers (성장동력)</a></li>
        <li><a href="#sec-risk">Risk Analysis (리스크 분석)</a></li>
        <li><a href="#sec-valuation">Valuation (기업가치 평가)</a></li>
        <li><a href="#sec-thesis">Investment Thesis (투자포인트)</a></li>
        <li><a href="#sec-scenario">Scenario Analysis (시나리오)</a></li>
        <li><a href="#sec-conclusion">Conclusion (결론)</a></li>
        <li><a href="#sec-refs">자료 출처</a></li>
      </ol>
    </nav>"""

    # ── 1. Executive Summary ──────────────────────────────────────
    if es:
        es_html = f"""
    <section id="sec-es" class="section">
      <h2 class="section-title">⚡ 1. Executive Summary (핵심 요약)</h2>
      <div class="two-col">
        <div class="card">
          <div class="card-label">기업 한 줄 정의</div>
          <div class="card-value">{_nl(es.get('one_line', '-'))}</div>
        </div>
        <div class="card">
          <div class="card-label">핵심 경쟁력</div>
          <div class="card-value">{_nl(es.get('key_competency', '-'))}</div>
        </div>
        <div class="card">
          <div class="card-label">현재 위치</div>
          <div class="card-value">{_nl(es.get('current_position', '-'))}</div>
        </div>
        <div class="card">
          <div class="card-label">향후 전망</div>
          <div class="card-value">{_nl(es.get('outlook', '-'))}</div>
        </div>
      </div>
      <div class="es-conclusion">📌 결론 : {_nl(es.get('conclusion', '-'))}</div>
    </section>"""
    else:
        es_html = _empty_section("sec-es", "⚡ 1. Executive Summary (핵심 요약)")

    # ── 2. Company Overview ───────────────────────────────────────
    overview_html = f"""
    <section id="sec-overview" class="section">
      <h2 class="section-title">📌 2. Company Overview (기업 개요)</h2>
      <table class="info-table">
        <tr><th>기업명</th><td>{co}</td></tr>
        <tr><th>티커 / 거래소</th><td>{ticker} / {exchange}</td></tr>
        <tr><th>설립연도</th><td>{ov.get('founded', '-')}</td></tr>
        <tr><th>본사 위치</th><td>{ov.get('hq', '-')}</td></tr>
        <tr><th>주요 산업</th><td>{ov.get('industry', '-')}</td></tr>
        <tr><th>임직원 수</th><td>{ov.get('employees', '-')}</td></tr>
      </table>
      <div class="text-box">{_nl(ov.get('summary', ''))}</div>
    </section>"""

    # ── 3. Industry Analysis ──────────────────────────────────────
    industry_html = f"""
    <section id="sec-industry" class="section">
      <h2 class="section-title">🌐 3. Industry Analysis (산업 분석)</h2>
      <table class="info-table">
        <tr><th>시장 규모 (TAM)</th><td>{mk.get('total_size', '-')}</td></tr>
        <tr><th>연평균 성장률</th><td>{mk.get('growth_rate', '-')}</td></tr>
        <tr><th>시장 점유율</th><td>{mk.get('market_share', '-')}</td></tr>
      </table>
      <div class="text-box">{_nl(mk.get('trend', ''))}</div>
    </section>"""

    # ── 4. Competitive Landscape ──────────────────────────────────
    if comp:
        def _comp_row(c_item):
            cls = ' class="highlight-row"' if c_item.get("is_target") else ""
            return (
                f'<tr{cls}>'
                f'<td class="left">{c_item.get("name","-")}</td>'
                f'<td>{c_item.get("revenue","-")}</td>'
                f'<td>{c_item.get("market_share","-")}</td>'
                f'<td class="left">{c_item.get("strategy","-")}</td></tr>'
            )
        comp_rows = "".join(_comp_row(c) for c in comp.get("competitors", [])) \
            or '<tr><td colspan="4" class="empty">데이터 없음</td></tr>'
        competitive_html = f"""
    <section id="sec-competitive" class="section">
      <h2 class="section-title">⚔️ 4. Competitive Landscape (경쟁 구도 분석)</h2>
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>기업명</th><th>매출</th><th>시장 점유율</th><th>핵심 전략</th></tr></thead>
          <tbody>{comp_rows}</tbody>
        </table>
      </div>
      <div class="sub-title">포지셔닝 분석</div>
      <div class="text-box">{_nl(comp.get('positioning', '-'))}</div>
      <div class="sub-title">강점 / 약점</div>
      <div class="two-col">
        <div class="card card-left"><div class="card-label">▲ 강점</div><div class="card-value">{_nl(comp.get('strengths', '-'))}</div></div>
        <div class="card card-left"><div class="card-label">▽ 약점</div><div class="card-value">{_nl(comp.get('weaknesses', '-'))}</div></div>
      </div>
    </section>"""
    else:
        competitive_html = _empty_section("sec-competitive", "⚔️ 4. Competitive Landscape (경쟁 구도 분석)")

    # ── 5. Business Model Analysis ────────────────────────────────
    if bm:
        bm_html = f"""
    <section id="sec-bm" class="section">
      <h2 class="section-title">⚙️ 5. Business Model Analysis (사업모델 분석)</h2>
      <div class="two-col">
        <div class="card card-left">
          <div class="card-label">💰 수익 구조</div>
          <div class="card-value">{_nl(bm.get('revenue_structure', '-'))}</div>
        </div>
        <div class="card card-left">
          <div class="card-label">📉 비용 구조</div>
          <div class="card-value">{_nl(bm.get('cost_structure', '-'))}</div>
        </div>
        <div class="card card-left">
          <div class="card-label">📈 수익성 구조</div>
          <div class="card-value">{_nl(bm.get('profitability', '-'))}</div>
        </div>
        <div class="card card-left">
          <div class="card-label">🏆 경쟁 우위</div>
          <div class="card-value">{_nl(bm.get('competitive_advantage', '-'))}</div>
        </div>
      </div>
    </section>"""
    else:
        bm_html = _empty_section("sec-bm", "⚙️ 5. Business Model Analysis (사업모델 분석)")

    # ── 6. Management Analysis ────────────────────────────────────
    exec_cards = "".join(
        f"""<div class="exec-card">
          <div class="exec-name">{e.get('name', '-')}</div>
          <div class="exec-title-tag">{e.get('title', '-')}</div>
          <div class="exec-detail">
            <span class="exec-label">학력</span> {e.get('education', '-')}<br>
            <span class="exec-label">이력</span> {_nl(e.get('career', '-'))}
          </div>
        </div>"""
        for e in exc
    ) or '<p class="empty">데이터 없음</p>'
    management_html = f"""
    <section id="sec-mgmt" class="section">
      <h2 class="section-title">👤 6. Management Analysis (경영진 분석)</h2>
      <div class="exec-grid">{exec_cards}</div>
    </section>"""

    # ── 7. Financial Analysis ─────────────────────────────────────
    fin_rows = "".join(
        f"""<tr>
          <td class="year-cell">{f.get('year', '-')}</td>
          <td>{f.get('revenue', '-')}</td>
          <td>{f.get('op_income', '-')}</td>
          <td>{f.get('net_income', '-')}</td>
          <td>{f.get('eps', '-')}</td>
          <td>{f.get('roe', '-')}</td>
        </tr>"""
        for f in fin
    ) or '<tr><td colspan="6" class="empty">데이터 없음</td></tr>'

    bs  = finx.get("balance_sheet", {})
    cf  = finx.get("cashflow", {})
    kpi = finx.get("kpi", {})

    bs_html = ""
    if bs:
        bs_html = f"""<div class="sub-title">재무상태표</div>
      <table class="info-table">
        <tr><th>총자산</th><td>{bs.get('assets', '-')}</td></tr>
        <tr><th>총부채</th><td>{bs.get('liabilities', '-')}</td></tr>
        <tr><th>자기자본</th><td>{bs.get('equity', '-')}</td></tr>
      </table>"""

    cf_html = ""
    if cf:
        cf_html = f"""<div class="sub-title">현금흐름</div>
      <table class="info-table">
        <tr><th>영업현금흐름</th><td>{cf.get('operating', '-')}</td></tr>
        <tr><th>투자현금흐름</th><td>{cf.get('investing', '-')}</td></tr>
        <tr><th>잉여현금흐름 (FCF)</th><td>{cf.get('fcf', '-')}</td></tr>
      </table>"""

    kpi_cards = "".join(
        f'<div class="kpi-card"><div class="kpi-label">{lbl}</div><div class="kpi-value">{kpi[key]}</div></div>'
        for lbl, key in [("ROE", "roe"), ("ROIC", "roic"), ("EBITDA", "ebitda"), ("영업이익률", "op_margin"), ("부채비율", "debt_ratio")]
        if kpi.get(key)
    )
    kpi_html = f'<div class="sub-title">주요 지표</div><div class="kpi-grid">{kpi_cards}</div>' if kpi_cards else ""

    financial_html = f"""
    <section id="sec-financial" class="section">
      <h2 class="section-title">💹 7. Financial Analysis (재무 분석)</h2>
      <div class="sub-title">손익계산서 (5개년)</div>
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>연도</th><th>매출액</th><th>영업이익</th><th>순이익</th><th>EPS</th><th>ROE</th></tr></thead>
          <tbody>{fin_rows}</tbody>
        </table>
      </div>
      {bs_html}{cf_html}{kpi_html}
    </section>"""

    # ── 8. Growth Drivers ─────────────────────────────────────────
    drivers_items = "".join(
        f"""<div class="driver-card">
          <div class="driver-title">🔹 {d.get('title', '')}</div>
          <div class="driver-detail">{_nl(d.get('detail', ''))}</div>
        </div>"""
        for d in drv
    ) or '<p class="empty">데이터 없음</p>'
    growth_html = f"""
    <section id="sec-growth" class="section">
      <h2 class="section-title">🚀 8. Growth Drivers (성장동력 분석)</h2>
      {drivers_items}
    </section>"""

    # ── 9. Risk Analysis ──────────────────────────────────────────
    if risk:
        risk_items = [
            ("사업 리스크",    "business"),
            ("기술 리스크",    "technology"),
            ("규제 리스크",    "regulatory"),
            ("경쟁 리스크",    "competitive"),
            ("거시경제 리스크", "macro"),
        ]
        risk_cards = "".join(
            f'<div class="risk-card"><div class="risk-card-title">⚠️ {lbl}</div>'
            f'<div class="risk-card-body">{_nl(risk.get(key, "-"))}</div></div>'
            for lbl, key in risk_items if risk.get(key)
        ) or '<p class="empty">데이터 없음</p>'
        risk_html = f"""
    <section id="sec-risk" class="section">
      <h2 class="section-title">⚠️ 9. Risk Analysis (리스크 분석)</h2>
      <div class="two-col">{risk_cards}</div>
    </section>"""
    else:
        risk_html = _empty_section("sec-risk", "⚠️ 9. Risk Analysis (리스크 분석)")

    # ── 10. Valuation ─────────────────────────────────────────────
    if val:
        is_listed = val.get("is_listed", True)
        rel_v = val.get("relative", {})
        abs_v = val.get("absolute", {})
        val_rows = ""
        if is_listed:
            for lbl, key in [("PER", "per"), ("PBR", "pbr"), ("EV/EBITDA", "ev_ebitda")]:
                if rel_v.get(key):
                    val_rows += f"<tr><th>{lbl}</th><td>{rel_v[key]}</td></tr>"
        if abs_v.get("dcf"):
            val_rows += f"<tr><th>DCF (절대가치)</th><td>{abs_v['dcf']}</td></tr>"
        if val.get("target_value"):
            val_rows += f"<tr><th>목표 기업가치</th><td>{val['target_value']}</td></tr>"
        if is_listed and val.get("target_price"):
            val_rows += f"<tr><th>적정 주가</th><td><strong>{val['target_price']}</strong></td></tr>"
        listed_tag = "📈 상장 기업" if is_listed else "🔒 비상장 기업 — 상대배수 일부 미적용"
        val_inner = f'<table class="info-table">{val_rows}</table>' if val_rows else '<p class="empty">데이터 없음</p>'
        valuation_html = f"""
    <section id="sec-valuation" class="section">
      <h2 class="section-title">💎 10. Valuation (기업가치 평가)</h2>
      <span class="listed-tag">{listed_tag}</span>
      {val_inner}
    </section>"""
    else:
        valuation_html = _empty_section("sec-valuation", "💎 10. Valuation (기업가치 평가)")

    # ── 11. Investment Thesis ─────────────────────────────────────
    if its:
        thesis_items = "".join(
            f"""<li class="thesis-item">
              <div class="thesis-num">{i + 1}</div>
              <div class="thesis-text">
                <div class="thesis-point">{t.get('point', '-')}</div>
                <div class="thesis-detail">{_nl(t.get('detail', ''))}</div>
              </div>
            </li>"""
            for i, t in enumerate(its)
        )
        thesis_html = f"""
    <section id="sec-thesis" class="section">
      <h2 class="section-title">🎯 11. Investment Thesis (투자포인트)</h2>
      <ul class="thesis-list">{thesis_items}</ul>
    </section>"""
    else:
        thesis_html = _empty_section("sec-thesis", "🎯 11. Investment Thesis (투자포인트)")

    # ── 12. Scenario Analysis ─────────────────────────────────────
    if scen:
        def _scen_card(key, css_cls, label_cls, label):
            s = scen.get(key, {})
            return (
                f'<div class="scenario-card {css_cls}">'
                f'<div class="scenario-label {label_cls}">{label}</div>'
                f'<div class="scenario-body">'
                f'<strong>조건 :</strong> {_nl(s.get("conditions", "-"))}<br><br>'
                f'<strong>전망 :</strong> {_nl(s.get("outlook", "-"))}'
                f'</div></div>'
            )
        scenario_html = f"""
    <section id="sec-scenario" class="section">
      <h2 class="section-title">🔮 12. Scenario Analysis (시나리오 분석)</h2>
      <div class="three-col">
        {_scen_card("bull", "scenario-bull", "label-bull", "🟢 Bull Case (낙관)")}
        {_scen_card("base", "scenario-base", "label-base", "🔵 Base Case (기준)")}
        {_scen_card("bear", "scenario-bear", "label-bear", "🔴 Bear Case (비관)")}
      </div>
    </section>"""
    else:
        scenario_html = _empty_section("sec-scenario", "🔮 12. Scenario Analysis (시나리오 분석)")

    # ── 13. Conclusion ────────────────────────────────────────────
    if conc:
        conclusion_html = f"""
    <section id="sec-conclusion" class="section">
      <h2 class="section-title">📝 13. Conclusion (결론)</h2>
      <div class="two-col">
        <div class="card">
          <div class="card-label">최종 평가</div>
          <div class="card-value">{_nl(conc.get('final_assessment', '-'))}</div>
        </div>
        <div class="card">
          <div class="card-label">핵심 체크포인트</div>
          <div class="card-value">{_nl(conc.get('checkpoints', '-'))}</div>
        </div>
      </div>
      <div class="text-box highlight-box" style="margin-top:14px;">
        <strong>향후 관전 포인트</strong><br>{_nl(conc.get('watchpoints', '-'))}
      </div>
    </section>"""
    else:
        conclusion_html = _empty_section("sec-conclusion", "📝 13. Conclusion (결론)")

    # ── 14. 자료 출처 ──────────────────────────────────────────────
    ref_items = "".join(
        f'<li><a href="{r.get("url","#")}" target="_blank">{r.get("title", r.get("url",""))}</a></li>'
        for r in refs
    ) or '<li class="empty">참고 URL 없음</li>'
    refs_html = f"""
    <section id="sec-refs" class="section">
      <h2 class="section-title">📎 14. 자료 출처</h2>
      <ul class="ref-list">{ref_items}</ul>
    </section>"""

    # ── 전체 조합 ─────────────────────────────────────────────────
    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title}</title>
  <style>{css}</style>
</head>
<body>
<div class="container">
  <header class="report-header">
    <div class="badge-concept">🔎 스드스(Sdeus) · {cm['label']}</div>
    <h1>{title}</h1>
    <div class="meta">
      작성일 : {report_date} &nbsp;|&nbsp;
      분석 대상 : {co} &nbsp;|&nbsp; 티커 : {ticker}
    </div>
  </header>
  {toc_html}
  {es_html}
  {overview_html}
  {industry_html}
  {competitive_html}
  {bm_html}
  {management_html}
  {financial_html}
  {growth_html}
  {risk_html}
  {valuation_html}
  {thesis_html}
  {scenario_html}
  {conclusion_html}
  {refs_html}
  <div class="report-footer">
    본 리포트는 스드스(Sdeus) 기업 분석 시스템에 의해 생성되었습니다.
  </div>
</div>
</body>
</html>"""


def save_report(data: dict) -> str:
    os.makedirs(REPORT_DIR, exist_ok=True)
    now = kst_now()
    company_slug = data.get("company_name", "unknown").replace(" ", "_")
    filename = f"{company_slug}_{now.strftime('%Y%m%d_%H%M')}.html"
    filepath = os.path.join(REPORT_DIR, filename)
    html = build_html_report(data)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[스드스] 리포트 저장 완료: {filepath}")
    return filepath
