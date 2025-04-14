import requests
import webbrowser
import http.server
import threading
import urllib.parse

# === Replace with your Azure app settings ===
client_id = "ce0c9aa4-6d58-4787-8dee-45285189899f"
client_secret = "TgJ8Q~ABdrKIHuWJ5x.ZR.96u53eJ5-vXugUoafQ"
redirect_uri = "http://localhost:8000"
scopes = ["Mail.ReadWrite"]

# === Step 1: OAuth Redirect Handler ===
class OAuthCallbackHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        self.server.auth_code = params.get("code", [None])[0]
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"<h1>You may close this window now.</h1>")

def start_server():
    server = http.server.HTTPServer(("localhost", 8000), OAuthCallbackHandler)
    server.handle_request()
    return server.auth_code

# === Step 2: Start OAuth Flow ===
auth_url = (
    "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?"
    f"client_id={client_id}&response_type=code&redirect_uri={redirect_uri}"
    f"&response_mode=query&scope={' '.join(scopes)}&state=12345"
)
print("Opening browser to authenticate...")
webbrowser.open(auth_url)

print("Waiting for authorization...")
auth_code = start_server()
print("Authorization code received!")

# === Step 3: Exchange Code for Token ===
token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
token_data = {
    "client_id": client_id,
    "client_secret": client_secret,
    "code": auth_code,
    "redirect_uri": redirect_uri,
    "grant_type": "authorization_code",
    "scope": " ".join(scopes),
}

token_response = requests.post(token_url, data=token_data).json()
access_token = token_response.get("access_token")

if not access_token:
    print("Failed to get access token:", token_response)
    exit()

print("Access token acquired!")

# === Step 4: Create Email Draft ===
headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

email_data = {
    "subject": "Draft Email from Joe AI",
    "body": {
        "contentType": "Text",
        "content": "Hi there,\n\nThis is a test draft email created using Microsoft Graph API and Joe AI!\n\nCheers,\nAhmed"
    },
    "toRecipients": []  # leave empty for now (user will fill manually later)
}

response = requests.post("https://graph.microsoft.com/v1.0/me/messages", headers=headers, json=email_data)

if response.status_code == 201:
    draft_id = response.json().get("id")
    print(f"Draft created successfully! Draft ID: {draft_id}")
else:
    print(f"Failed to create draft: {response.status_code}")
    print(response.text)
