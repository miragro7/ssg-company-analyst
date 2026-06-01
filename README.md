# ssg-company-analyst

AI-powered company analysis skill for Claude Code.
Type the trigger below in your Telegram chat — Claude automatically
researches, builds a 14-section report, and delivers HTML + Word files.

**Trigger**
> 스드스야 [기업명] 분석해줘

Examples
- 스드스야 삼성전자 분석해줘
- 스드스야 Coforge 분석해줘
- 스드스야 LG CNS 조사해줘

---

## What it does

When triggered, the skill:

1. **Researches** the company using WebSearch + WebFetch across IR materials, news, and financial data
2. **Builds a 14-section report** covering everything from Executive Summary to Investment Thesis
3. **Generates two files** — an HTML report and a Word (.docx) document
4. **Delivers both files** via Telegram automatically

### Report Sections (14 fixed)

| # | Section |
|---|---------|
| 1 | Executive Summary (핵심 요약) |
| 2 | Company Overview (기업 개요) |
| 3 | Industry Analysis (산업 분석) |
| 4 | Competitive Landscape (경쟁 구도) |
| 5 | Business Model Analysis (사업모델) |
| 6 | Management Analysis (경영진) |
| 7 | Financial Analysis (재무 분석) |
| 8 | Growth Drivers (성장동력) |
| 9 | Risk Analysis (리스크 분석) |
| 10 | Valuation (기업가치 평가) |
| 11 | Investment Thesis (투자포인트) |
| 12 | Scenario Analysis (시나리오) |
| 13 | Conclusion (결론) |
| 14 | 자료 출처 |

---

## Installation

### 1. Clone this repository

```bash
git clone https://github.com/miragro7/ssg-company-analyst.git
cd ssg-company-analyst
```

### 2. Install Python dependencies

```bash
pip install pyTelegramBotAPI python-docx lxml
```

### 3. Configure environment variables

```bash
cp .env.example .env
# Edit .env and fill in your BOT_TOKEN and CHAT_ID
```

### 4. Add skill instruction to your CLAUDE.md

Copy the contents of `skill.md` into your project's `CLAUDE.md`.

---

## Configuration

Create a `.env` file (or set environment variables directly):

```env
BOT_TOKEN=your_telegram_bot_token_here
CHAT_ID=your_telegram_chat_id_here
REPORT_DIR=./report
```

| Variable | Description |
|----------|-------------|
| `BOT_TOKEN` | Telegram Bot API token (from @BotFather) |
| `CHAT_ID` | Telegram chat ID to deliver reports |
| `REPORT_DIR` | Directory to save generated reports (default: `./report`) |

---

## File Structure

```
ssg-company-analyst/
├── sdeus_pipeline.py        # Main pipeline: HTML + Word + Telegram delivery
├── sdeus_reporter.py        # HTML report generator (14-section template)
├── skill.md                 # Claude Code skill definition (add to CLAUDE.md)
├── .env.example             # Environment variable template
├── requirements.txt         # Python dependencies
├── report/                  # Generated reports saved here (gitignored)
└── skills/
    └── sds-word-writer/
        └── scripts/
            └── generate.py  # Word (.docx) document generator
```

---

## Usage

Once configured, use the following triggers in Claude Code or your Telegram bot:

```
스드스야 삼성전자 분석해줘
스드스야 Apple 분석해줘
스드스야 NVIDIA 조사해줘
```

Claude will:
- Automatically detect the company name
- Research across IR filings, news, and financial databases
- Generate the 14-section report
- Save as `{company}_{YYYYMMDD_HHMM}.html` and `.docx`
- Send both files to your Telegram

---

## Report Concept Types

Add a concept keyword before the company name for focused analysis:

| Concept | Focus |
|---------|-------|
| `투자판단` | Investment decision |
| `재무건전성` | Financial health |
| `성장전략` | Growth strategy |
| `사업성과리스크` | Business performance & risk |
| *(default)* | 종합분석 — Comprehensive analysis |

Example: `스드스야 투자판단 삼성전자 분석해줘`

---

## License

MIT License — free to use and modify.

Built for [Claude Code](https://claude.ai/code) by Mir Kim.
