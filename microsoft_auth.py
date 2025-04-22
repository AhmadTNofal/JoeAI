import requests, webbrowser, http.server, urllib.parse

# Microsoft application credentials (used for OAuth)
client_id = "ce0c9aa4-6d58-4787-8dee-45285189899f"
client_secret = "TgJ8Q~ABdrKIHuWJ5x.ZR.96u53eJ5-vXugUoafQ"  
redirect_uri = "http://localhost:8000"  # Local redirect for receiving auth code
scopes = ["Tasks.ReadWrite", "Mail.ReadWrite"]  # Permissions for Microsoft To Do and Outlook

# Custom HTTP handler for OAuth callback to capture authorization code
class OAuthCallbackHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        # Parse the URL and extract the authorization code
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        self.server.auth_code = params.get("code", [None])[0]

        # Respond to browser indicating completion
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"<h1>You may close this window.</h1>")

# Function to get the access token using OAuth 2.0 Authorization Code flow
def get_access_token():
    # Step 1: Generate and open the Microsoft login URL for user authentication
    auth_url = (
        f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize?"
        f"client_id={client_id}&response_type=code&redirect_uri={redirect_uri}"
        f"&response_mode=query&scope={' '.join(scopes)}&state=12345"
    )
    webbrowser.open(auth_url)

    # Step 2: Launch a temporary local HTTP server to receive the auth code
    server = http.server.HTTPServer(("localhost", 8000), OAuthCallbackHandler)
    server.handle_request()  # Blocks until callback is received
    code = server.auth_code  # Retrieved from the request

    # Step 3: Exchange the authorization code for an access token
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    token_data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
        "code": code,
        "scope": " ".join(scopes)
    }

    # Send POST request to obtain token and return the access token from response
    response = requests.post(token_url, data=token_data).json()
    return response.get("access_token")
