#!/usr/bin/env python3
"""
Outlook Draft Creator
Creates Outlook drafts from Excel contact list using Microsoft Graph API
"""

import argparse
import csv
import json
import os
import time
from pathlib import Path
from typing import Dict, List, Optional

import msal
import pandas as pd
import requests
from dotenv import load_dotenv
from jinja2 import Template

# === USE EXACTLY; DO NOT CHANGE WORDING ===
SUBJECT_TEXT = """PASTE YOUR SUBJECT HERE (you may include {{placeholders}})"""

BODY_TEXT = """PASTE YOUR BODY HERE EXACTLY.
You may reference Excel fields like {{first_name}}, {{company}}, {{role}}, {{observation}}.
If you don't want placeholders, leave plain text.
"""
# ==========================================

# Microsoft Graph API configuration
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Mail.ReadWrite", "offline_access"]
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"

# Token cache file
TOKEN_CACHE_FILE = "token_cache.bin"


class OutlookDraftCreator:
    def __init__(self, client_id: str, from_name: str, from_email: str):
        self.client_id = client_id
        self.from_name = from_name
        self.from_email = from_email
        self.app = msal.PublicClientApplication(
            client_id,
            authority=AUTHORITY,
            token_cache=msal.SerializableTokenCache()
        )
        self.access_token = None
        self._load_token_cache()
    
    def _load_token_cache(self):
        """Load token cache from file if it exists"""
        if os.path.exists(TOKEN_CACHE_FILE):
            with open(TOKEN_CACHE_FILE, "r") as f:
                cache_data = f.read()
                self.app.token_cache.deserialize(cache_data)
    
    def _save_token_cache(self):
        """Save token cache to file"""
        cache_data = self.app.token_cache.serialize()
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache_data)
    
    def authenticate(self) -> bool:
        """Authenticate using device code flow"""
        # Check if we have a valid cached token
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(SCOPES, account=accounts[0])
            if result:
                self.access_token = result["access_token"]
                self._save_token_cache()
                return True
        
        # Use device code flow for new authentication
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            print("Failed to create device flow")
            return False
        
        print(f"To sign in, use a web browser to open the page {flow['verification_uri']}")
        print(f"Enter the code {flow['user_code']} to authenticate.")
        
        result = self.app.acquire_token_by_device_flow(flow)
        if "access_token" in result:
            self.access_token = result["access_token"]
            self._save_token_cache()
            return True
        
        print(f"Authentication failed: {result.get('error_description', 'Unknown error')}")
        return False
    
    def create_draft(self, email: str, subject: str, body: str) -> Dict:
        """Create a draft email using Microsoft Graph API"""
        if not self.access_token:
            return {"status": "error", "error": "No access token"}
        
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": email
                    }
                }
            ]
        }
        
        try:
            response = requests.post(
                f"{GRAPH_ENDPOINT}/me/messages",
                headers=headers,
                json=payload
            )
            
            if response.status_code == 201:
                result = response.json()
                return {
                    "status": "success",
                    "draft_id": result.get("id", "unknown"),
                    "error": None
                }
            else:
                return {
                    "status": "error",
                    "error": f"HTTP {response.status_code}: {response.text}"
                }
        
        except Exception as e:
            return {
                "status": "error",
                "error": str(e)
            }
    
    def personalize_with_ai(self, row_data: Dict) -> str:
        """Add AI personalization if API key is available"""
        openai_api_key = os.getenv("OPENAI_API_KEY")
        if not openai_api_key:
            return ""
        
        # Simple prompt for personalization
        prompt = f"""Based on this contact info, write a brief, personalized sentence (max 35 words) to add to a cold email:
        Name: {row_data.get('first_name', '')} {row_data.get('last_name', '')}
        Company: {row_data.get('company', '')}
        Role: {row_data.get('role', '')}
        Observation: {row_data.get('observation', '')}
        
        Write a natural, genuine-sounding personalization:"""
        
        try:
            headers = {
                "Authorization": f"Bearer {openai_api_key}",
                "Content-Type": "application/json"
            }
            
            payload = {
                "model": "gpt-3.5-turbo",
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": 50
            }
            
            response = requests.post(
                "https://api.openai.com/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=10
            )
            
            if response.status_code == 200:
                result = response.json()
                personalization = result["choices"][0]["message"]["content"].strip()
                # Ensure it's not too long
                if len(personalization) > 200:
                    personalization = personalization[:197] + "..."
                return f"\n\n{personalization}"
        
        except Exception as e:
            print(f"AI personalization failed: {e}")
        
        return ""


def render_template(template_text: str, row_data: Dict) -> str:
    """Render Jinja2 template with row data"""
    try:
        template = Template(template_text)
        return template.render(**row_data)
    except Exception as e:
        print(f"Template rendering error: {e}")
        return template_text


def log_result(csv_file: str, email: str, company: str, first_name: str, 
               subject: str, draft_id: str, status: str):
    """Log result to CSV file"""
    file_exists = os.path.exists(csv_file)
    
    with open(csv_file, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        
        if not file_exists:
            writer.writerow(['email', 'company', 'first_name', 'subject', 'draft_id', 'status'])
        
        writer.writerow([email, company, first_name, subject, draft_id, status])


def main():
    parser = argparse.ArgumentParser(description="Create Outlook drafts from Excel contact list")
    parser.add_argument("--excel", required=True, help="Excel file path")
    parser.add_argument("--sheet", required=True, help="Sheet name")
    parser.add_argument("--from_name", required=True, help="Sender name")
    parser.add_argument("--from_email", required=True, help="Sender email")
    parser.add_argument("--delay_ms", type=int, default=1000, help="Delay between drafts in milliseconds")
    parser.add_argument("--log_csv", required=True, help="CSV log file path")
    parser.add_argument("--ai-personalize", action="store_true", help="Enable AI personalization")
    parser.add_argument("--client-id", required=True, help="Azure AD App Client ID")
    
    args = parser.parse_args()
    
    # Load environment variables
    load_dotenv()
    
    # Validate Excel file
    if not os.path.exists(args.excel):
        print(f"Error: Excel file '{args.excel}' not found")
        return
    
    # Read Excel data
    try:
        df = pd.read_excel(args.excel, sheet_name=args.sheet)
        print(f"Loaded {len(df)} contacts from Excel")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return
    
    # Initialize Outlook creator
    creator = OutlookDraftCreator(args.client_id, args.from_name, args.from_email)
    
    # Authenticate
    if not creator.authenticate():
        print("Authentication failed. Exiting.")
        return
    
    print("Authentication successful!")
    
    # Process each contact
    for index, row in df.iterrows():
        row_data = row.to_dict()
        
        # Clean up data - replace NaN with empty strings
        for key, value in row_data.items():
            if pd.isna(value):
                row_data[key] = ""
        
        # Render templates
        subject = render_template(SUBJECT_TEXT, row_data)
        body = render_template(BODY_TEXT, row_data)
        
        # Add AI personalization if enabled
        if args.ai_personalize:
            personalization = creator.personalize_with_ai(row_data)
            if personalization:
                body += personalization
        
        # Create draft
        email = row_data.get('email', '')
        if not email:
            print(f"Row {index + 1}: No email address, skipping")
            continue
        
        print(f"Creating draft for {email}...")
        result = creator.create_draft(email, subject, body)
        
        # Log result
        log_result(
            args.log_csv,
            email,
            row_data.get('company', ''),
            row_data.get('first_name', ''),
            subject,
            result.get('draft_id', ''),
            result['status']
        )
        
        if result['status'] == 'success':
            print(f"✓ Draft created successfully (ID: {result['draft_id']})")
        else:
            print(f"✗ Draft creation failed: {result.get('error', 'Unknown error')}")
        
        # Rate limiting
        if index < len(df) - 1:  # Don't delay after the last one
            delay_seconds = args.delay_ms / 1000.0
            print(f"Waiting {delay_seconds:.1f} seconds...")
            time.sleep(delay_seconds)
    
    print(f"\nProcessing complete! Check '{args.log_csv}' for results.")


if __name__ == "__main__":
    main() 