import requests, webbrowser, http.server, urllib.parse
import os
from dotenv import load_dotenv

load_dotenv()

# === Microsoft Azure App Registration Details ===
client_id = os.getenv("CLIENT_ID")                          # Application (client) ID from Azure
client_secret = os.getenv("CLIENT_SECRET")                  # Secret generated for the app (keep this safe)
redirect_uri = os.getenv("REDIRECT_URI")                    # Redirect URI used to capture the auth code locally
scopes = os.getenv("SCOPES").split()                        # Permissions requested for Microsoft Graph API access

# === HTTP Server Handler to Capture Redirect with Auth Code ===
class OAuthCallbackHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        """
        Handles GET requests sent to the redirect URI.
        Extracts the 'code' parameter from the query string after successful user authentication.
        """
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        self.server.auth_code = params.get("code", [None])[0]  # Store the received auth code on the server object
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"<h1>You may close this window.</h1>")  # Notify user auth succeeded

# === Function to Perform OAuth Flow and Fetch Access Token ===
def get_access_token():
    """
    Launches browser for Microsoft sign-in, handles redirect via localhost,
    exchanges auth code for an access token, and returns the token string.
    """
    # Construct the authorization URL with required query parameters
    auth_url = (
        f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize?"
        f"client_id={client_id}&response_type=code&redirect_uri={redirect_uri}"
        f"&response_mode=query&scope={' '.join(scopes)}&state=12345"
    )

    # Open the user's default browser to complete the sign-in
    webbrowser.open(auth_url)

    # Start a local HTTP server to capture the redirect and receive the auth code
    server = http.server.HTTPServer(("localhost", 8000), OAuthCallbackHandler)
    server.handle_request()
    code = server.auth_code  # Auth code received from redirect

    # Define token exchange request parameters
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    token_data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
        "code": code,
        "scope": " ".join(scopes)
    }

    # Perform token exchange request
    response = requests.post(token_url, data=token_data).json()
    return response.get("access_token")

# === Function to Create Outlook Draft Email ===
def create_email_draft(access_token, subject, body_content):
    """
    Uses Microsoft Graph API to create a draft email in the authenticated user's mailbox.

    Args:
        access_token (str): OAuth bearer token for Graph API.
        subject (str): Subject of the email.
        body_content (str): Body text of the email.

    Returns:
        Tuple[bool, str]: True and draft ID if successful, otherwise False and None.
    """
    url = "https://graph.microsoft.com/v1.0/me/messages"  # Graph API endpoint to create draft message
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    data = {
        "subject": subject,
        "body": {
            "contentType": "Text",  # Content type set to plain text
            "content": body_content
        }
    }

    # Make the API request to create the draft
    response = requests.post(url, headers=headers, json=data)
    if response.status_code == 201:
        return True, response.json().get("id")  # Return success and the draft ID
    else:
        print(" Failed to create draft:", response.status_code, response.text)
        return False, None
