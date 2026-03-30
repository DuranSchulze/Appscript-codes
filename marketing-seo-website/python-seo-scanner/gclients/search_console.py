"""Google Search Console API client."""

from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from datetime import datetime, timedelta
from typing import List, Dict, Optional

import config
from gclients.auth import get_search_console_credentials


class SearchConsoleClient:
    def __init__(self):
        self.credentials = None
        self.service = None
        self.property_url = config.GSC_PROPERTY
    
    def connect(self) -> bool:
        """Connect to Search Console API."""
        try:
            self.credentials = get_search_console_credentials()
            self.service = build('searchconsole', 'v1', credentials=self.credentials)
            return True
        except Exception as e:
            print(f"Failed to connect to Search Console: {e}")
            return False
    
    def get_search_analytics(
        self,
        page_url: str,
        days: int = 28
    ) -> Optional[Dict]:
        """Get search analytics for a specific page."""
        if not self.service or not self.property_url:
            return None
        
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
        
        try:
            request = {
                'startDate': start_date,
                'endDate': end_date,
                'dimensions': ['page', 'query'],
                'dimensionFilterGroups': [{
                    'filters': [{
                        'dimension': 'page',
                        'operator': 'equals',
                        'expression': page_url
                    }]
                }],
                'rowLimit': 10
            }
            
            response = self.service.searchanalytics().query(
                siteUrl=self.property_url,
                body=request
            ).execute()
            
            rows = response.get('rows', [])
            if not rows:
                return None
            
            total_clicks = sum(row.get('clicks', 0) for row in rows)
            total_impressions = sum(row.get('impressions', 0) for row in rows)
            avg_ctr = sum(row.get('ctr', 0) for row in rows) / len(rows) if rows else 0
            avg_position = sum(row.get('position', 0) for row in rows) / len(rows) if rows else 0
            
            top_queries = [
                {
                    'query': row['keys'][1],
                    'clicks': row.get('clicks', 0),
                    'impressions': row.get('impressions', 0),
                    'ctr': row.get('ctr', 0),
                    'position': row.get('position', 0)
                }
                for row in rows[:5]
            ]
            
            return {
                'clicks': total_clicks,
                'impressions': total_impressions,
                'ctr': round(avg_ctr, 4),
                'avg_position': round(avg_position, 1),
                'top_queries': top_queries
            }
            
        except Exception as e:
            print(f"Error fetching GSC data for {page_url}: {e}")
            return None
    
    def get_site_summary(self, days: int = 28) -> Optional[Dict]:
        """Get overall site search analytics summary."""
        if not self.service or not self.property_url:
            return None
        
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
        
        try:
            request = {
                'startDate': start_date,
                'endDate': end_date,
                'dimensions': [],
                'rowLimit': 1
            }
            
            response = self.service.searchanalytics().query(
                siteUrl=self.property_url,
                body=request
            ).execute()
            
            rows = response.get('rows', [])
            if rows:
                row = rows[0]
                return {
                    'clicks': row.get('clicks', 0),
                    'impressions': row.get('impressions', 0),
                    'ctr': row.get('ctr', 0),
                    'avg_position': row.get('position', 0)
                }
            
            return None
            
        except Exception as e:
            print(f"Error fetching site summary: {e}")
            return None
