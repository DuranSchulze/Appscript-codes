"""Configuration management for SEO Scanner."""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).parent
CREDENTIALS_DIR = BASE_DIR / "credentials"
OUTPUT_DIR = BASE_DIR / "reports"

CREDENTIALS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
DEFAULT_AI_PROVIDER = os.getenv("DEFAULT_AI_PROVIDER", "gemini")

GSC_PROPERTY = os.getenv("GSC_PROPERTY", "")
GA4_PROPERTY_ID = os.getenv("GA4_PROPERTY_ID", "")

# Google Ads API Configuration
GOOGLE_ADS_DEVELOPER_TOKEN = os.getenv("GOOGLE_ADS_DEVELOPER_TOKEN", "")
GOOGLE_ADS_LOGIN_CUSTOMER_ID = os.getenv("GOOGLE_ADS_LOGIN_CUSTOMER_ID", "")
GOOGLE_ADS_CUSTOMER_ID = os.getenv("GOOGLE_ADS_CUSTOMER_ID", "")

# SEMrush API Configuration
SEMRUSH_API_KEY = os.getenv("SEMRUSH_API_KEY", "")

DEFAULT_MAX_PAGES = int(os.getenv("DEFAULT_MAX_PAGES", "100"))
DEFAULT_MAX_DEPTH = int(os.getenv("DEFAULT_MAX_DEPTH", "2"))
DEFAULT_CRAWL_DELAY = float(os.getenv("DEFAULT_CRAWL_DELAY", "0.5"))

GOOGLE_CLIENT_SECRET_FILE = CREDENTIALS_DIR / "client_secret.json"
GSC_TOKEN_FILE = CREDENTIALS_DIR / "gsc_token.json"
GA4_TOKEN_FILE = CREDENTIALS_DIR / "ga4_token.json"
GADS_TOKEN_FILE = CREDENTIALS_DIR / "gads_token.json"

GSC_SCOPES = ["https://www.googleapis.com/auth/webmasters.readonly"]
GA4_SCOPES = ["https://www.googleapis.com/auth/analytics.readonly"]
GADS_SCOPES = ["https://www.googleapis.com/auth/adwords"]

USER_AGENT = "Mozilla/5.0 (compatible; SEOScannerBot/1.0; +https://github.com/seo-scanner)"
