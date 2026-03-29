"""Gemini API client for AI keyword recommendations."""

import time
from typing import Dict, Optional, List
import requests

import config


class GeminiClient:
    def __init__(self, api_key: str = None, model: str = "gemini-1.5-flash"):
        self.api_key = api_key or config.GEMINI_API_KEY
        self.model = model
        self.base_url = "https://generativelanguage.googleapis.com/v1beta/models"
        
        if not self.api_key:
            raise ValueError("Gemini API key is required. Set GEMINI_API_KEY in .env file.")
    
    def generate_content(self, prompt: str, max_retries: int = 3) -> Optional[str]:
        """Generate content using Gemini API."""
        url = f"{self.base_url}/{self.model}:generateContent?key={self.api_key}"
        
        payload = {
            "contents": [{
                "parts": [{"text": prompt}]
            }],
            "generationConfig": {
                "temperature": 0.7,
                "maxOutputTokens": 1024,
            }
        }
        
        for attempt in range(max_retries):
            try:
                response = requests.post(
                    url,
                    json=payload,
                    headers={"Content-Type": "application/json"},
                    timeout=60
                )
                response.raise_for_status()
                
                data = response.json()
                
                if "candidates" in data and data["candidates"]:
                    content = data["candidates"][0].get("content", {})
                    parts = content.get("parts", [])
                    if parts:
                        return parts[0].get("text", "")
                
                return None
                
            except requests.exceptions.RequestException as e:
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
                    continue
                print(f"Gemini API error after {max_retries} retries: {e}")
                return None
        
        return None
    
    def test_connection(self) -> bool:
        """Test API connection with a simple prompt."""
        try:
            result = self.generate_content("Say 'Connection successful'", max_retries=1)
            return result is not None
        except Exception:
            return False


def format_keyword_prompt(
    page_title: str,
    meta_description: str,
    h1: str,
    content_snippet: str,
    current_queries: List[Dict] = None
) -> str:
    """Format prompt for keyword recommendations."""
    
    queries_text = ""
    if current_queries:
        queries_text = "\nCurrent Google Search Console queries driving traffic to this page:\n"
        for q in current_queries[:5]:
            queries_text += f"- '{q['query']}': {q['clicks']} clicks, {q['impressions']} impressions, position {q['position']}\n"
    
    prompt = f"""You are an expert SEO strategist. Analyze this webpage and provide keyword recommendations.

Page Title: {page_title}
Meta Description: {meta_description}
H1 Heading: {h1}
Content Snippet: {content_snippet}{queries_text}

Provide your analysis in this exact format:

PRIMARY KEYWORD: [single most important target keyword - 2-4 words, high search intent]

SECONDARY KEYWORDS:
- [keyword 1]
- [keyword 2]
- [keyword 3]
- [keyword 4]

CONTENT GAPS: [What content is missing that would help this page rank better? 1-2 sentences]

SEARCH INTENT: [Informational/Commercial/Transactional/Navigational]

OPTIMIZATION TIP: [One specific, actionable tip to improve this page's SEO in under 50 words]

Be concise and specific. Avoid generic advice."""
    
    return prompt
