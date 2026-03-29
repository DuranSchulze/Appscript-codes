"""AI-powered keyword recommendation engine."""

from typing import Dict, Optional, List
import re

from ai.gemini_client import GeminiClient, format_keyword_prompt as gemini_prompt
from ai.openai_client import OpenAIClient, format_keyword_prompt as openai_prompt
import config


class KeywordRecommender:
    def __init__(self, provider: str = None):
        self.provider = provider or config.DEFAULT_AI_PROVIDER
        self.client = None
        
        if self.provider == "gemini":
            self.client = GeminiClient()
        elif self.provider == "openai":
            self.client = OpenAIClient()
        else:
            raise ValueError(f"Unknown AI provider: {self.provider}")
    
    def test_connection(self) -> bool:
        """Test the AI provider connection."""
        if self.client:
            return self.client.test_connection()
        return False
    
    def get_recommendations(
        self,
        page_title: str,
        meta_description: str,
        h1: str,
        content_snippet: str,
        current_queries: List[Dict] = None
    ) -> Dict[str, str]:
        """Get AI keyword recommendations for a page."""
        
        if not self.client:
            return {"error": "AI client not initialized"}
        
        if self.provider == "gemini":
            prompt = gemini_prompt(page_title, meta_description, h1, content_snippet, current_queries)
        else:
            prompt = openai_prompt(page_title, meta_description, h1, content_snippet, current_queries)
        
        try:
            response = self.client.generate_content(prompt)
            
            if not response:
                return {"error": "No response from AI provider"}
            
            return self._parse_response(response)
            
        except Exception as e:
            return {"error": str(e)}
    
    def _parse_response(self, response: str) -> Dict[str, str]:
        """Parse AI response into structured data."""
        result = {
            "primary_keyword": "",
            "secondary_keywords": [],
            "content_gaps": "",
            "search_intent": "",
            "optimization_tip": "",
            "raw_response": response
        }
        
        lines = response.split('\n')
        current_section = None
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            if line.startswith('PRIMARY KEYWORD:'):
                result["primary_keyword"] = line.split(':', 1)[1].strip()
                current_section = None
            elif line.startswith('SECONDARY KEYWORDS:'):
                current_section = "secondary"
            elif line.startswith('CONTENT GAPS:'):
                result["content_gaps"] = line.split(':', 1)[1].strip()
                current_section = None
            elif line.startswith('SEARCH INTENT:'):
                result["search_intent"] = line.split(':', 1)[1].strip()
                current_section = None
            elif line.startswith('OPTIMIZATION TIP:'):
                result["optimization_tip"] = line.split(':', 1)[1].strip()
                current_section = None
            elif current_section == "secondary" and line.startswith('-'):
                keyword = line[1:].strip()
                if keyword:
                    result["secondary_keywords"].append(keyword)
        
        return result


def parse_keyword_recommendations(recommendations: Dict) -> str:
    """Format recommendations for display."""
    if "error" in recommendations:
        return f"Error: {recommendations['error']}"
    
    lines = []
    
    if recommendations.get("primary_keyword"):
        lines.append(f"Primary: {recommendations['primary_keyword']}")
    
    if recommendations.get("secondary_keywords"):
        lines.append(f"Secondary: {', '.join(recommendations['secondary_keywords'])}")
    
    if recommendations.get("search_intent"):
        lines.append(f"Intent: {recommendations['search_intent']}")
    
    if recommendations.get("optimization_tip"):
        lines.append(f"Tip: {recommendations['optimization_tip']}")
    
    return " | ".join(lines) if lines else "No recommendations available"
