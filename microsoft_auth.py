import requests, webbrowser, http.server, urllib.parse

client_id = "ce0c9aa4-6d58-4787-8dee-45285189899f"
client_secret = "TgJ8Q~ABdrKIHuWJ5x.ZR.96u53eJ5-vXugUoafQ"
redirect_uri = "http://localhost:8000"
scopes = ["Tasks.ReadWrite"]

class OAuthCallbackHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        self.server.auth_code = params.get("code", [None])[0]
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"<h1>You may close this window.</h1>")

def get_access_token():

    auth_url = (
        f"https://login.microsoftonline.com/common/oauth2/v2.0/authorize?"
        f"client_id={client_id}&response_type=code&redirect_uri={redirect_uri}"
        f"&response_mode=query&scope={' '.join(scopes)}&state=12345"
    )
    webbrowser.open(auth_url)
    server = http.server.HTTPServer(("localhost", 8000), OAuthCallbackHandler)
    server.handle_request()
    code = server.auth_code

    # Step 2: Token exchange
    token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    token_data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
        "code": code,
        "scope": " ".join(scopes)
    }

    response = requests.post(token_url, data=token_data).json()
    return response.get("access_token")
