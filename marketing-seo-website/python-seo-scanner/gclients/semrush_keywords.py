"""SEMrush API client for keyword metrics."""

import requests
from typing import List, Dict, Optional
from urllib.parse import quote

import config


class SEMrushKeywordClient:
    """Client for fetching keyword metrics from SEMrush API."""
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key or config.SEMRUSH_API_KEY
        self.base_url = "https://api.semrush.com"
        
        if not self.api_key:
            raise ValueError("SEMrush API key is required. Set SEMRUSH_API_KEY in .env file.")
    
    def get_keyword_metrics(
        self,
        keywords: List[str],
        database: str = "us",  # Default to US database
    ) -> List[Dict]:
        """
        Get keyword metrics from SEMrush.
        
        Args:
            keywords: List of keyword phrases
            database: Country database code (us, uk, ca, au, etc.)
        
        Returns:
            List of dicts with keyword metrics
        """
        if not keywords:
            return []
        
        try:
            # SEMrush API allows batch requests
            # Build the request
            keyword_param = ";".join(keywords[:100])  # SEMrush allows up to 100 keywords
            
            params = {
                "type": "phrase_all",  # Get all metrics for keywords
                "key": self.api_key,
                "phrase": keyword_param,
                "database": database,
                "export_columns": "Ph,Nq,Cp,Co,Nr,Td",
                # Ph = Phrase, Nq = Search Volume, Cp = CPC, Co = Competition, Nr = Number of Results, Td = Trends
            }
            
            response = requests.get(
                f"{self.base_url}/",
                params=params,
                timeout=60
            )
            response.raise_for_status()
            
            # Parse SEMrush CSV-like response
            results = self._parse_response(response.text)
            
            return results
            
        except requests.exceptions.RequestException as e:
            print(f"SEMrush API error: {e}")
            return []
        except Exception as e:
            print(f"Error fetching SEMrush keyword metrics: {e}")
            return []
    
    def _parse_response(self, response_text: str) -> List[Dict]:
        """Parse SEMrush CSV response into list of dicts."""
        lines = response_text.strip().split('\n')
        
        if len(lines) < 2:
            return []
        
        # First line is header
        headers = lines[0].split(';')
        
        results = []
        for line in lines[1:]:
            if not line.strip():
                continue
            
            values = line.split(';')
            
            # Map SEMrush columns to our format
            # Ph;Nq;Cp;Co;Nr;Td
            result = {
                'keyword': values[0] if len(values) > 0 else '',
                'search_volume': int(values[1]) if len(values) > 1 and values[1].isdigit() else 0,
                'cpc': float(values[2]) if len(values) > 2 and values[2] else 0,
                'competition': float(values[3]) if len(values) > 3 and values[3] else 0,
                'number_of_results': int(values[4]) if len(values) > 4 and values[4].isdigit() else 0,
                'trends': values[5] if len(values) > 5 else '',
                'source': 'semrush'
            }
            
            # Convert competition to readable format (0-1 scale to descriptive)
            comp = result['competition']
            if comp < 0.33:
                result['competition_level'] = "LOW"
            elif comp < 0.66:
                result['competition_level'] = "MEDIUM"
            else:
                result['competition_level'] = "HIGH"
            
            results.append(result)
        
        return results
    
    def get_keyword_suggestions(
        self,
        seed_keyword: str,
        database: str = "us",
        limit: int = 10
    ) -> List[Dict]:
        """
        Get keyword suggestions based on a seed keyword.
        
        Returns related keywords with their metrics.
        """
        try:
            params = {
                "type": "phrase_related",
                "key": self.api_key,
                "phrase": seed_keyword,
                "database": database,
                "display_limit": limit,
                "export_columns": "Ph,Nq,Cp,Co",
            }
            
            response = requests.get(
                f"{self.base_url}/",
                params=params,
                timeout=60
            )
            response.raise_for_status()
            
            return self._parse_response(response.text)
            
        except Exception as e:
            print(f"Error fetching SEMrush keyword suggestions: {e}")
            return []
    
    def test_connection(self) -> bool:
        """Test API connection with a simple query."""
        try:
            results = self.get_keyword_metrics(["seo"], database="us")
            return len(results) > 0
        except Exception:
            return False


def format_semrush_metrics_table(metrics: List[Dict]) -> str:
    """Format SEMrush keyword metrics as a readable table string."""
    if not metrics:
        return "No keyword metrics available"
    
    lines = []
    lines.append("| Keyword | Search Volume | CPC | Competition | Results |")
    lines.append("|---------|--------------|-----|-------------|----------|")
    
    for m in metrics:
        keyword = m['keyword'][:35] + "..." if len(m['keyword']) > 35 else m['keyword']
        volume = f"{m['search_volume']:,}"
        cpc = f"${m['cpc']:.2f}" if m['cpc'] > 0 else "-"
        comp = m.get('competition_level', 'UNKNOWN')
        results = f"{m['number_of_results']:,}" if m['number_of_results'] > 0 else "-"
        
        lines.append(f"| {keyword} | {volume} | {cpc} | {comp} | {results} |")
    
    return "\n".join(lines)
