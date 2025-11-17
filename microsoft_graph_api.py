"""Microsoft Graph API Integration for Outlook, OneDrive, and SharePoint

This module provides a Python class to interact with Microsoft Graph API
for accessing user's Outlook emails, OneDrive files, and SharePoint sites.

Requires: msal, requests, python-dotenv
"""

import msal
import requests
import os
from dotenv import load_dotenv

load_dotenv()

# Azure App Credentials
CLIENT_ID = os.getenv('MICROSOFT_CLIENT_ID', '92149bb7-b052-4d4b-9d39-6cb1440716db')
CLIENT_SECRET = os.getenv('MICROSOFT_CLIENT_SECRET', '')
TENANT_ID = os.getenv('MICROSOFT_TENANT_ID', '2224abd7-7085-434d-b50a-add325728a07')
REDIRECT_URI = os.getenv('MICROSOFT_REDIRECT_URI', 'http://localhost:8000/oauth/callback/microsoft')

# Microsoft Graph API endpoint
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'

# Define scopes (permissions)
SCOPES = [
    'User.Read',
    'Mail.Read',
    'Mail.Send',
    'Files.Read.All',
    'Sites.Read.All'
]


class MicrosoftGraphAPI:
    """Microsoft Graph API Client for Outlook, OneDrive, SharePoint"""
    
    def __init__(self, client_id=None, client_secret=None, tenant_id=None, redirect_uri=None):
        self.client_id = client_id or CLIENT_ID
        self.client_secret = client_secret or CLIENT_SECRET
        self.tenant_id = tenant_id or TENANT_ID
        self.redirect_uri = redirect_uri or REDIRECT_URI
        self.token = None
        
    def get_auth_url(self):
        """Step 1: Get authorization URL for user to login"""
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            client_credential=self.client_secret
        )
        
        auth_url = app.get_authorization_request_url(
            scopes=SCOPES,
            redirect_uri=self.redirect_uri
        )
        return auth_url
    
    def get_token_from_code(self, auth_code):
        """Step 2: Exchange authorization code for access token"""
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            client_credential=self.client_secret
        )
        
        result = app.acquire_token_by_authorization_code(
            code=auth_code,
            scopes=SCOPES,
            redirect_uri=self.redirect_uri
        )
        
        if "access_token" in result:
            self.token = result['access_token']
            return result
        else:
            raise Exception(f"Error getting token: {result.get('error_description')}")
    
    def make_api_call(self, endpoint, method='GET', data=None):
        """Make API call to Microsoft Graph"""
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        
        url = f"{GRAPH_API_ENDPOINT}/{endpoint}"
        
        if method == 'GET':
            response = requests.get(url, headers=headers)
        elif method == 'POST':
            response = requests.post(url, headers=headers, json=data)
        
        return response.json()
    
    # === OUTLOOK EMAIL METHODS ===
    
    def get_user_profile(self):
        """Get user's profile information"""
        return self.make_api_call('me')
    
    def get_emails(self, top=10):
        """Get user's emails from Outlook"""
        return self.make_api_call(f'me/messages?$top={top}')
    
    def get_inbox_emails(self, top=10):
        """Get emails from inbox folder"""
        return self.make_api_call(f'me/mailFolders/inbox/messages?$top={top}')
    
    def search_emails(self, query):
        """Search emails by subject or content"""
        return self.make_api_call(f'me/messages?$search="{query}"')
    
    def send_email(self, to_email, subject, body):
        """Send an email via Outlook"""
        email_data = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": to_email
                        }
                    }
                ]
            }
        }
        return self.make_api_call('me/sendMail', method='POST', data=email_data)
    
    # === ONEDRIVE METHODS ===
    
    def get_onedrive_root(self):
        """Get root folder contents of OneDrive"""
        return self.make_api_call('me/drive/root/children')
    
    def get_file_content(self, file_id):
        """Download file content from OneDrive"""
        return self.make_api_call(f'me/drive/items/{file_id}/content')
    
    def search_onedrive(self, query):
        """Search files in OneDrive"""
        return self.make_api_call(f'me/drive/root/search(q="{query}")')
    
    def get_recent_files(self):
        """Get recently accessed files"""
        return self.make_api_call('me/drive/recent')
    
    # === SHAREPOINT METHODS ===
    
    def get_sharepoint_sites(self):
        """Get SharePoint sites user has access to"""
        return self.make_api_call('sites?search=*')
    
    def get_site_lists(self, site_id):
        """Get lists from a SharePoint site"""
        return self.make_api_call(f'sites/{site_id}/lists')
    
    def get_site_drives(self, site_id):
        """Get document libraries from SharePoint site"""
        return self.make_api_call(f'sites/{site_id}/drives')
