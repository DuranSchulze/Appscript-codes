"""Utility helper functions."""

from datetime import datetime
from typing import Optional
from urllib.parse import urlparse


def format_datetime(dt: datetime) -> str:
    """Format datetime for display."""
    return dt.strftime('%Y-%m-%d %H:%M:%S')


def truncate_text(text: str, max_length: int = 100) -> str:
    """Truncate text to max length."""
    if not text:
        return ""
    if len(text) <= max_length:
        return text
    return text[:max_length - 3] + "..."


def get_page_path(url: str, base_url: str) -> str:
    """Extract page path from URL."""
    parsed = urlparse(url)
    return parsed.path or "/"


def score_seo_health(
    has_title: bool,
    has_meta_description: bool,
    has_h1: bool,
    title_length: int,
    meta_length: int,
    status_code: int
) -> tuple:
    """Score SEO health and return (score, status, issues)."""
    score = 100
    issues = []
    
    if not has_title:
        score -= 25
        issues.append("Missing title tag")
    elif title_length < 20:
        score -= 10
        issues.append(f"Title too short ({title_length} chars)")
    elif title_length > 60:
        score -= 10
        issues.append(f"Title too long ({title_length} chars)")
    
    if not has_meta_description:
        score -= 20
        issues.append("Missing meta description")
    elif meta_length < 50:
        score -= 5
        issues.append(f"Meta description too short ({meta_length} chars)")
    elif meta_length > 160:
        score -= 5
        issues.append(f"Meta description too long ({meta_length} chars)")
    
    if not has_h1:
        score -= 15
        issues.append("Missing H1 heading")
    
    if status_code != 200:
        score -= 30
        issues.append(f"HTTP {status_code} error")
    
    if score >= 90:
        status = "Good"
    elif score >= 70:
        status = "Needs Improvement"
    else:
        status = "Poor"
    
    return score, status, issues
