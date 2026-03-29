"""OpenAI API client for AI keyword recommendations."""

import time
from typing import Dict, Optional, List
from openai import OpenAI

import config


class OpenAIClient:
    def __init__(self, api_key: str = None, model: str = "gpt-4o-mini"):
        self.api_key = api_key or config.OPENAI_API_KEY
        self.model = model
        
        if not self.api_key:
            raise ValueError("OpenAI API key is required. Set OPENAI_API_KEY in .env file.")
        
        self.client = OpenAI(api_key=self.api_key)
    
    def generate_content(self, prompt: str, max_retries: int = 3) -> Optional[str]:
        """Generate content using OpenAI API."""
        
        for attempt in range(max_retries):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": "You are an expert SEO strategist. Provide concise, actionable keyword recommendations."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.7,
                    max_tokens=1024
                )
                
                return response.choices[0].message.content
                
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)
                    continue
                print(f"OpenAI API error after {max_retries} retries: {e}")
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
    
    prompt = f"""Analyze this webpage and provide keyword recommendations.

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

Be concise and specific."""
    
    return prompt
