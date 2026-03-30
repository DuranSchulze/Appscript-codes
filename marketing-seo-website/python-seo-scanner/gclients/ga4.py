"""Google Analytics 4 (GA4) API client."""

from googleapiclient.discovery import build
from datetime import datetime, timedelta
from typing import Dict, Optional

import config
from gclients.auth import get_ga4_credentials


class GA4Client:
    def __init__(self):
        self.credentials = None
        self.service = None
        self.property_id = config.GA4_PROPERTY_ID
    
    def connect(self) -> bool:
        """Connect to GA4 API."""
        try:
            self.credentials = get_ga4_credentials()
            self.service = build('analyticsdata', 'v1beta', credentials=self.credentials)
            return True
        except Exception as e:
            print(f"Failed to connect to GA4: {e}")
            return False
    
    def get_page_metrics(
        self,
        page_path: str,
        days: int = 28
    ) -> Optional[Dict]:
        """Get metrics for a specific page."""
        if not self.service or not self.property_id:
            return None
        
        property_name = f"properties/{self.property_id}"
        
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
        
        try:
            request = {
                'dateRanges': [{'startDate': start_date, 'endDate': end_date}],
                'metrics': [
                    {'name': 'sessions'},
                    {'name': 'totalUsers'},
                    {'name': 'newUsers'},
                    {'name': 'screenPageViews'},
                    {'name': 'averageEngagementTime'},
                    {'name': 'engagementRate'},
                ],
                'dimensionFilter': {
                    'filter': {
                        'fieldName': 'pagePath',
                        'stringFilter': {
                            'matchType': 'EXACT',
                            'value': page_path
                        }
                    }
                }
            }
            
            response = self.service.properties().runReport(
                property=property_name,
                body=request
            ).execute()
            
            rows = response.get('rows', [])
            if not rows:
                return None
            
            row = rows[0]
            metrics = row.get('metricValues', [])
            
            return {
                'sessions': int(metrics[0].get('value', 0)) if len(metrics) > 0 else 0,
                'users': int(metrics[1].get('value', 0)) if len(metrics) > 1 else 0,
                'new_users': int(metrics[2].get('value', 0)) if len(metrics) > 2 else 0,
                'page_views': int(metrics[3].get('value', 0)) if len(metrics) > 3 else 0,
                'avg_engagement_time': round(float(metrics[4].get('value', 0)), 2) if len(metrics) > 4 else 0,
                'engagement_rate': round(float(metrics[5].get('value', 0)), 4) if len(metrics) > 5 else 0,
            }
            
        except Exception as e:
            print(f"Error fetching GA4 data for {page_path}: {e}")
            return None
    
    def get_site_summary(self, days: int = 28) -> Optional[Dict]:
        """Get overall site GA4 summary."""
        if not self.service or not self.property_id:
            return None
        
        property_name = f"properties/{self.property_id}"
        
        end_date = datetime.now().strftime('%Y-%m-%d')
        start_date = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')
        
        try:
            request = {
                'dateRanges': [{'startDate': start_date, 'endDate': end_date}],
                'metrics': [
                    {'name': 'sessions'},
                    {'name': 'totalUsers'},
                    {'name': 'newUsers'},
                    {'name': 'screenPageViews'},
                ]
            }
            
            response = self.service.properties().runReport(
                property=property_name,
                body=request
            ).execute()
            
            rows = response.get('rows', [])
            if rows:
                metrics = rows[0].get('metricValues', [])
                return {
                    'sessions': int(metrics[0].get('value', 0)) if len(metrics) > 0 else 0,
                    'users': int(metrics[1].get('value', 0)) if len(metrics) > 1 else 0,
                    'new_users': int(metrics[2].get('value', 0)) if len(metrics) > 2 else 0,
                    'page_views': int(metrics[3].get('value', 0)) if len(metrics) > 3 else 0,
                }
            
            return None
            
        except Exception as e:
            print(f"Error fetching GA4 site summary: {e}")
            return None
