import requests
import json
from datetime import datetime, timedelta

class OutlookEmailSearcher:
    def __init__(self, access_token):
        self.access_token = access_token
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
    
    def search_emails(self, search_string, folder="inbox", max_results=50):
        """
        Search emails in Outlook using Microsoft Graph API
        
        Args:
            search_string (str): Text to search for in email body/subject
            folder (str): Folder to search in (default: "inbox")
            max_results (int): Maximum number of results to return
        """
        # Build the search query
        search_query = f"\"{search_string}\""
        
        # API endpoint for searching messages
        endpoint = f"{self.base_url}/me/mailFolders/{folder}/messages"
        
        params = {
            '$search': search_query,
            '$top': max_results,
            '$select': 'subject,receivedDateTime,from,bodyPreview,id',
            '$orderby': 'receivedDateTime desc'
        }
        
        try:
            response = requests.get(endpoint, headers=self.headers, params=params)
            response.raise_for_status()
            
            data = response.json()
            messages = data.get('value', [])
            
            if not messages:
                print(f"No emails found containing '{search_string}'")
                return []
            
            print(f"Found {len(messages)} emails containing '{search_string}':\n")
            
            for message in messages:
                subject = message.get('subject', 'No Subject')
                received = message.get('receivedDateTime', 'Unknown Date')
                sender = message.get('from', {}).get('emailAddress', {}).get('name', 'Unknown Sender')
                body_preview = message.get('bodyPreview', 'No preview available')
                
                print(f"Subject: {subject}")
                print(f"From: {sender}")
                print(f"Received: {received}")
                print(f"Preview: {body_preview[:100]}...")
                print("-" * 50)
            
            return messages
            
        except requests.exceptions.RequestException as e:
            print(f"Error searching emails: {e}")
            return []
    
    def get_email_content(self, message_id):
        """Get full email content by message ID"""
        endpoint = f"{self.base_url}/me/messages/{message_id}"
        params = {'$select': 'subject,body,from,receivedDateTime'}
        
        try:
            response = requests.get(endpoint, headers=self.headers, params=params)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"Error getting email content: {e}")
            return None

def main():
    print("Outlook Email Search Tool for macOS")
    print("=" * 40)
    
    # You'll need to get an access token from Microsoft Graph API
    # This requires setting up an Azure app registration
    access_token = input("Enter your Microsoft Graph API access token: ").strip()
    
    if not access_token:
        print("Access token is required. Please set up Microsoft Graph API authentication.")
        print("\nTo get an access token:")
        print("1. Register an app in Azure Portal")
        print("2. Grant Mail.Read permissions")
        print("3. Use the Microsoft Authentication Library (MSAL) to get tokens")
        return
    
    searcher = OutlookEmailSearcher(access_token)
    
    while True:
        search_term = input("\nEnter search term (or 'quit' to exit): ").strip()
        
        if search_term.lower() == 'quit':
            break
        
        if not search_term:
            print("Please enter a search term.")
            continue
        
        searcher.search_emails(search_term)

if __name__ == "__main__":
    main() 