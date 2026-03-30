"""Keyword metrics integration - combines GSC keywords with Google Ads data."""

from typing import Dict, List, Optional, Set
from collections import defaultdict

from gclients.search_console import SearchConsoleClient
from gclients.ads_keywords import GoogleAdsKeywordClient


class KeywordMetricsIntegration:
    """Integrates Search Console keywords with Google Ads metrics."""
    
    def __init__(
        self,
        search_console_client: Optional[SearchConsoleClient] = None,
        ads_client: Optional[GoogleAdsKeywordClient] = None
    ):
        self.gsc_client = search_console_client
        self.ads_client = ads_client
        self.keyword_data: Dict[str, Dict] = {}
    
    def extract_keywords_from_gsc(
        self,
        days: int = 28,
        min_impressions: int = 10,
        limit: int = 100
    ) -> List[Dict]:
        """
        Extract unique keywords from Search Console data.
        
        Returns list of keywords with their aggregated metrics across all pages.
        """
        if not self.gsc_client or not self.gsc_client.service:
            return []
        
        from datetime import datetime, timedelta
        
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
        
        try:
            request = {
                'startDate': start_date,
                'endDate': end_date,
                'dimensions': ['query'],
                'rowLimit': limit
            }
            
            response = self.gsc_client.service.searchanalytics().query(
                siteUrl=self.gsc_client.property_url,
                body=request
            ).execute()
            
            keywords = []
            for row in response.get('rows', []):
                impressions = row.get('impressions', 0)
                if impressions >= min_impressions:
                    keywords.append({
                        'keyword': row['keys'][0],
                        'clicks': row.get('clicks', 0),
                        'impressions': impressions,
                        'ctr': row.get('ctr', 0),
                        'avg_position': row.get('position', 0)
                    })
            
            return keywords
            
        except Exception as e:
            print(f"Error extracting keywords from GSC: {e}")
            return []
    
    def enrich_keywords_with_ads_metrics(
        self,
        keywords: List[Dict],
        location_id: str = "2840",
        language_id: str = "1000"
    ) -> List[Dict]:
        """
        Enrich GSC keywords with Google Ads metrics.
        
        Returns keywords with search volume, competition, and bid data.
        """
        if not self.ads_client or not self.ads_client.client:
            return keywords
        
        keyword_list = [k['keyword'] for k in keywords]
        
        try:
            ads_metrics = self.ads_client.get_keyword_metrics(
                keywords=keyword_list,
                location_id=location_id,
                language_id=language_id
            )
            
            # Create lookup by keyword
            ads_lookup = {m['keyword']: m for m in ads_metrics}
            
            # Merge data
            enriched = []
            for k in keywords:
                keyword_text = k['keyword']
                merged = {
                    'keyword': keyword_text,
                    'gsc_clicks': k['clicks'],
                    'gsc_impressions': k['impressions'],
                    'gsc_ctr': k['ctr'],
                    'gsc_avg_position': k['avg_position'],
                    'search_volume': 0,
                    'competition': 'UNKNOWN',
                    'competition_index': 0,
                    'suggested_bid_low': 0,
                    'suggested_bid_high': 0,
                    'opportunity_score': 0
                }
                
                if keyword_text in ads_lookup:
                    ads = ads_lookup[keyword_text]
                    merged['search_volume'] = ads['avg_monthly_searches']
                    merged['competition'] = ads['competition']
                    merged['competition_index'] = ads['competition_index']
                    merged['suggested_bid_low'] = ads['low_top_of_page_bid']
                    merged['suggested_bid_high'] = ads['high_top_of_page_bid']
                    
                    # Calculate opportunity score
                    # High impressions in GSC + low position = high opportunity
                    # High search volume + low competition = high opportunity
                    position_factor = max(0, 20 - merged['gsc_avg_position']) / 20  # 0-1
                    volume_factor = min(1, merged['search_volume'] / 10000)  # 0-1
                    competition_factor = (100 - merged['competition_index']) / 100  # 0-1
                    
                    merged['opportunity_score'] = round(
                        (position_factor * 0.4 + volume_factor * 0.3 + competition_factor * 0.3) * 100
                    )
                
                enriched.append(merged)
            
            # Sort by opportunity score
            enriched.sort(key=lambda x: x['opportunity_score'], reverse=True)
            
            return enriched
            
        except Exception as e:
            print(f"Error enriching keywords with Ads metrics: {e}")
            return keywords
    
    def get_top_opportunities(
        self,
        days: int = 28,
        top_n: int = 20,
        location_id: str = "2840",
        language_id: str = "1000"
    ) -> List[Dict]:
        """
        Get top keyword opportunities from GSC + Ads data.
        
        Returns ranked list of keywords with best opportunity to improve.
        """
        keywords = self.extract_keywords_from_gsc(days=days)
        
        if not keywords:
            return []
        
        if self.ads_client and self.ads_client.client:
            keywords = self.enrich_keywords_with_ads_metrics(
                keywords,
                location_id=location_id,
                language_id=language_id
            )
        
        return keywords[:top_n]


def format_keyword_opportunities_table(keywords: List[Dict]) -> str:
    """Format keyword opportunities as a Markdown table."""
    if not keywords:
        return "No keyword data available."
    
    has_ads_data = any(k.get('search_volume', 0) > 0 for k in keywords)
    
    if has_ads_data:
        lines = [
            "| Keyword | GSC Position | GSC Clicks | Search Volume | Competition | Bid Range | Opportunity |",
            "|---------|-------------|------------|---------------|-------------|-----------|-------------|"
        ]
        
        for k in keywords[:20]:
            keyword = k['keyword'][:30] + "..." if len(k['keyword']) > 30 else k['keyword']
            position = f"{k['gsc_avg_position']:.1f}"
            clicks = f"{k['gsc_clicks']:,}"
            volume = f"{k.get('search_volume', 0):,}" if k.get('search_volume', 0) > 0 else "-"
            comp = k.get('competition', 'N/A')[:4]
            bid = f"${k.get('suggested_bid_low', 0):.2f}-${k.get('suggested_bid_high', 0):.2f}" if k.get('suggested_bid_low', 0) > 0 else "-"
            opp = f"{k.get('opportunity_score', 0)}/100"
            
            lines.append(f"| {keyword} | {position} | {clicks} | {volume} | {comp} | {bid} | {opp} |")
    else:
        lines = [
            "| Keyword | Position | Clicks | Impressions | CTR |",
            "|---------|----------|--------|-------------|-----|"
        ]
        
        for k in keywords[:20]:
            keyword = k['keyword'][:40] + "..." if len(k['keyword']) > 40 else k['keyword']
            position = f"{k['gsc_avg_position']:.1f}"
            clicks = f"{k['gsc_clicks']:,}"
            impressions = f"{k['gsc_impressions']:,}"
            ctr = f"{k['gsc_ctr']:.2%}"
            
            lines.append(f"| {keyword} | {position} | {clicks} | {impressions} | {ctr} |")
    
    return "\n".join(lines)


def get_quick_wins(keywords: List[Dict], min_position: float = 15, max_position: float = 5) -> List[Dict]:
    """
    Identify 'quick win' keywords - ranking between positions 5-15 with decent volume.
    
    These are keywords where small improvements can yield significant traffic gains.
    """
    quick_wins = []
    
    for k in keywords:
        position = k.get('gsc_avg_position', 100)
        impressions = k.get('gsc_impressions', 0)
        
        if max_position <= position <= min_position and impressions > 100:
            quick_wins.append(k)
    
    return quick_wins
