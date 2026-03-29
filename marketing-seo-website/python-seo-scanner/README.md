# SEO Scanner

A Python CLI tool that crawls websites, analyzes SEO, fetches real Google Search Console and GA4 data, and generates AI-powered keyword recommendations—all output to a readable Markdown report.

## Features

- **Website Crawling**: Discovers pages via sitemap or recursive crawling
- **On-Page SEO Analysis**: Titles, meta descriptions, H1-H3, canonical, robots, word count
- **Google Search Console Integration**: Real impressions, clicks, CTR, rankings
- **Google Analytics 4 Integration**: Sessions, users, engagement metrics
- **AI Keyword Recommendations**: Gemini or OpenAI-powered keyword suggestions per page
- **Google Ads Keyword Metrics**: Real search volume, competition, and suggested bid data
- **Markdown Reports**: Beautiful, readable reports with priority action items

## Quick Start

### 1. Install Dependencies

```bash
cd marketing-seo-website/python-seo-scanner
pip install -r requirements.txt
```

### 2. Configure Environment

```bash
cp .env.example .env
# Edit .env with your API keys
```

### 3. Run Basic Scan

```bash
python seo_scanner.py scan https://example.com
```

### 4. Full Scan with Google APIs and AI

```bash
python seo_scanner.py scan https://example.com --gsc --ga4 --ai-provider gemini
```

## Setup

### AI Provider Setup

**Gemini (Google)**:

1. Get API key: https://ai.google.dev/
2. Add to `.env`: `GEMINI_API_KEY=your_key`

**OpenAI**:

1. Get API key: https://platform.openai.com/
2. Add to `.env`: `OPENAI_API_KEY=your_key`

### Google API Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project or select existing
3. Enable APIs:
   - Search Console API
   - Analytics Data API
   - Google Ads API (for keyword metrics)
4. Create OAuth2 credentials (Desktop application)
5. Download JSON and save as `credentials/client_secret.json`
6. Run: `python seo_scanner.py auth-setup`

### Google Ads API Setup (for Keyword Metrics)

1. Go to [Google Ads API Center](https://ads.google.com/aw/apicenter)
2. Apply for developer token (takes 1-2 business days for approval)
3. Add to `.env`:
   ```env
   GOOGLE_ADS_DEVELOPER_TOKEN=your_token_here
   GOOGLE_ADS_CUSTOMER_ID=your_customer_id  # Optional, will auto-detect if not set
   ```
4. Ensure you're using the same Google account that has access to a Google Ads account

### SEMrush API Setup (Alternative for Keyword Metrics)

1. Get SEMrush API access at https://www.semrush.com/api-management/
2. Add to `.env`:
   ```env
   SEMRUSH_API_KEY=your_api_key_here
   ```
3. SEMrush provides search volume, CPC, and competition data without needing OAuth

### Search Console & GA4 Configuration

Add to `.env`:

```env
GSC_PROPERTY=https://example.com/
GA4_PROPERTY_ID=123456789
```

## Usage

### Basic Scan

```bash
python seo_scanner.py scan https://example.com
```

### Get Keyword Metrics

```bash
# Get keyword metrics via Google Ads (default)
python seo_scanner.py keyword-metrics "seo services" "digital marketing" "web design"

# Get keyword metrics via SEMrush
python seo_scanner.py keyword-metrics "seo services" "digital marketing" "web design" --provider semrush

# Get keyword metrics from both providers
python seo_scanner.py keyword-metrics "seo services" --provider both

# Keyword metrics with custom location (UK = 2826)
python seo_scanner.py keyword-metrics "seo agency" --location 2826 --output keywords.md
```

### Scan with Keyword Metrics Enrichment

```bash
python seo_scanner.py scan https://example.com \
  --max-pages 50 \
  --max-depth 3 \
  --delay 1.0 \
  --gsc \
  --ga4 \
  --keyword-metrics \
  --keyword-provider semrush \
  --ai-provider openai \
  --output ./my-report.md
```

### Commands

| Command                         | Description                          |
| ------------------------------- | ------------------------------------ |
| `scan URL`                      | Main scan command                    |
| `auth-setup`                    | Authenticate with Google APIs        |
| `clear-auth`                    | Clear stored Google tokens           |
| `test-ai [provider]`            | Test AI connection                   |
| `keyword-metrics [keywords...]` | Get search volume & competition data |

### Scan Options

| Option                    | Description                                      |
| ------------------------- | ------------------------------------------------ |
| `--max-pages, -p`         | Maximum pages to crawl (default: 100)            |
| `--max-depth, -d`         | Maximum crawl depth (default: 2)                 |
| `--delay`                 | Delay between requests in seconds (default: 0.5) |
| `--gsc`                   | Connect to Google Search Console                 |
| `--ga4`                   | Connect to Google Analytics 4                    |
| `--keyword-metrics, -km`  | Enrich GSC keywords with search volume data      |
| `--keyword-provider, -kp` | Provider: `google-ads`, `semrush`, or `both`     |
| `--ai-provider`           | AI provider: `gemini` or `openai`                |
| `--no-ai`                 | Skip AI recommendations                          |
| `--output, -o`            | Custom output file path                          |

## Report Output

Reports are saved as Markdown files with:

- **Executive Summary**: Pages analyzed, issues count, Google data status
- **Priority Actions**: Critical and medium priority fixes ranked by impact
- **Page Analysis**: Per-page SEO audit with:
  - On-page SEO elements
  - Google Search Console metrics (if connected)
  - Google Analytics 4 metrics (if connected)
  - AI keyword recommendations (if enabled)

## Project Structure

```
python-seo-scanner/
├── seo_scanner.py           # Main CLI entry point
├── config.py                # Configuration management
├── requirements.txt         # Python dependencies
├── .env.example             # Environment template
├── credentials/             # OAuth credentials (gitignored)
├── crawler/
│   ├── site_crawler.py      # Website crawling
│   └── page_parser.py       # HTML parsing
├── google/
│   ├── auth.py              # OAuth2 authentication
│   ├── search_console.py    # GSC API client
│   └── ga4.py               # GA4 API client
├── ai/
│   ├── gemini_client.py     # Gemini API wrapper
│   ├── openai_client.py     # OpenAI API wrapper
│   └── keyword_recommender.py # AI recommendations
├── reporter/
│   └── markdown_reporter.py # Report generation
└── utils/
    └── helpers.py           # Utility functions
```

## Requirements

- Python 3.8+
- Internet connection
- API keys (for AI and/or Google features)

## License

MIT
