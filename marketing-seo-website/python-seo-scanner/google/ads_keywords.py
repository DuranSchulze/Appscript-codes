"""Google Ads API client for keyword metrics."""

from google.ads.googleads.client import GoogleAdsClient
from google.ads.googleads.errors import GoogleAdsException
from typing import List, Dict, Optional
import os

import config
from google.auth import load_credentials


class GoogleAdsKeywordClient:
    """Client for fetching keyword metrics from Google Ads API."""
    
    def __init__(self, developer_token: str = None, login_customer_id: str = None):
        self.developer_token = developer_token or config.GOOGLE_ADS_DEVELOPER_TOKEN
        self.login_customer_id = login_customer_id or config.GOOGLE_ADS_LOGIN_CUSTOMER_ID
        self.client = None
        self.customer_id = config.GOOGLE_ADS_CUSTOMER_ID
    
    def connect(self) -> bool:
        """Connect to Google Ads API."""
        if not self.developer_token:
            print("Google Ads Developer Token not configured")
            return False
        
        try:
            # Load OAuth credentials
            if not os.path.exists(config.GADS_TOKEN_FILE):
                credentials = load_credentials(str(config.GADS_TOKEN_FILE), config.GADS_SCOPES)
            else:
                credentials = load_credentials(str(config.GADS_TOKEN_FILE), config.GADS_SCOPES)
            
            # Initialize Google Ads client configuration
            googleads_client = GoogleAdsClient.load_from_dict({
                'developer_token': self.developer_token,
                'login_customer_id': self.login_customer_id if self.login_customer_id else None,
                'oauth2_client_id': credentials.client_id,
                'oauth2_client_secret': credentials.client_secret,
                'oauth2_refresh_token': credentials.refresh_token,
                'use_proto_plus': True,
            })
            
            self.client = googleads_client
            
            # Test connection by getting accessible customers
            customer_service = self.client.get_service("CustomerService")
            accessible_customers = customer_service.list_accessible_customers()
            
            if accessible_customers.resource_names:
                # If no specific customer ID set, use the first accessible one
                if not self.customer_id:
                    self.customer_id = accessible_customers.resource_names[0].split('/')[-1]
                print(f"Connected to Google Ads account: {self.customer_id}")
                return True
            else:
                print("No accessible Google Ads accounts found")
                return False
                
        except Exception as e:
            print(f"Failed to connect to Google Ads API: {e}")
            return False
    
    def get_keyword_metrics(
        self,
        keywords: List[str],
        location_id: str = "2840",  # USA by default
        language_id: str = "1000"    # English by default
    ) -> List[Dict]:
        """Get search volume and metrics for keywords."""
        if not self.client or not self.customer_id:
            return []
        
        try:
            keyword_plan_idea_service = self.client.get_service("KeywordPlanIdeaService")
            
            # Build request
            request = self.client.get_type("GenerateKeywordIdeasRequest")
            request.customer_id = self.customer_id
            
            # Set keyword seed
            keyword_seed = self.client.get_type("KeywordSeed")
            for keyword in keywords[:20]:  # Limit to 20 keywords per request
                keyword_seed.keywords.append(keyword)
            request.keyword_seed = keyword_seed
            
            # Set geo target
            request.geo_target_constants.append(
                f"geoTargetConstants/{location_id}"
            )
            
            # Set language
            request.language = f"languageConstants/{language_id}"
            
            # Set network
            request.keyword_plan_network = (
                self.client.enums.KeywordPlanNetworkEnum.GOOGLE_SEARCH
            )
            
            # Execute request
            response = keyword_plan_idea_service.generate_keyword_ideas(request=request)
            
            results = []
            for idea in response:
                metrics = idea.keyword_idea_metrics
                
                # Convert micros to actual currency
                low_bid = metrics.low_top_of_page_bid_micros / 1000000 if metrics.low_top_of_page_bid_micros else 0
                high_bid = metrics.high_top_of_page_bid_micros / 1000000 if metrics.high_top_of_page_bid_micros else 0
                
                results.append({
                    'keyword': idea.text,
                    'avg_monthly_searches': metrics.avg_monthly_searches,
                    'competition': self._format_competition(metrics.competition),
                    'competition_index': metrics.competition_index,
                    'low_top_of_page_bid': low_bid,
                    'high_top_of_page_bid': high_bid,
                })
            
            return results
            
        except GoogleAdsException as ex:
            for error in ex.failure.errors:
                print(f"Google Ads API error: {error.message}")
            return []
        except Exception as e:
            print(f"Error fetching keyword metrics: {e}")
            return []
    
    def _format_competition(self, competition_enum) -> str:
        """Convert competition enum to readable string."""
        competition_map = {
            0: "UNSPECIFIED",
            1: "UNKNOWN", 
            2: "LOW",
            3: "MEDIUM",
            4: "HIGH"
        }
        return competition_map.get(competition_enum, "UNKNOWN")


def format_keyword_metrics_table(metrics: List[Dict]) -> str:
    """Format keyword metrics as a readable table string."""
    if not metrics:
        return "No keyword metrics available"
    
    lines = []
    lines.append("| Keyword | Avg Monthly Searches | Competition | Suggested Bid |")
    lines.append("|---------|---------------------|-------------|---------------|")
    
    for m in metrics:
        keyword = m['keyword'][:35] + "..." if len(m['keyword']) > 35 else m['keyword']
        searches = f"{m['avg_monthly_searches']:,}"
        comp = m['competition']
        bid = f"${m['low_top_of_page_bid']:.2f}-${m['high_top_of_page_bid']:.2f}"
        
        lines.append(f"| {keyword} | {searches} | {comp} | {bid} |")
    
    return "\n".join(lines)
