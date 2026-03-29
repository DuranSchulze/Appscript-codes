"""Google OAuth2 authentication handler."""

import os
import json
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.exceptions import RefreshError

import config


def load_credentials(token_file: str, scopes: list):
    """Load credentials from token file or run OAuth flow."""
    creds = None
    
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, scopes)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError:
                creds = None
        
        if not creds:
            if not os.path.exists(config.GOOGLE_CLIENT_SECRET_FILE):
                raise FileNotFoundError(
                    f"Google client secret file not found: {config.GOOGLE_CLIENT_SECRET_FILE}\n"
                    "Please download your OAuth2 credentials from Google Cloud Console "
                    "and save them to the credentials/ folder."
                )
            
            flow = InstalledAppFlow.from_client_secrets_file(
                config.GOOGLE_CLIENT_SECRET_FILE, scopes
            )
            creds = flow.run_local_server(port=0)
        
        with open(token_file, 'w') as token:
            token.write(creds.to_json())
    
    return creds


def get_search_console_credentials():
    """Get credentials for Search Console API."""
    return load_credentials(str(config.GSC_TOKEN_FILE), config.GSC_SCOPES)


def get_ga4_credentials():
    """Get credentials for GA4 API."""
    return load_credentials(str(config.GA4_TOKEN_FILE), config.GA4_SCOPES)


def clear_credentials():
    """Clear all stored credentials for re-authentication."""
    for token_file in [config.GSC_TOKEN_FILE, config.GA4_TOKEN_FILE, config.GADS_TOKEN_FILE]:
        if os.path.exists(token_file):
            os.remove(token_file)
            print(f"Removed: {token_file}")
