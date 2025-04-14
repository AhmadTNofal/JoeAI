import requests
import webbrowser
import http.server
import urllib.parse

# 1. Azure app settings (set these up in Azure Portal)
client_id = "ce0c9aa4-6d58-4787-8dee-45285189899f"
client_secret = "TgJ8Q~ABdrKIHuWJ5x.ZR.96u53eJ5-vXugUoafQ"  
redirect_uri = "http://localhost:8000"
scopes = ["Tasks.ReadWrite"]

# 2. Start a local web server to catch the redirect with the code
class OAuthCallbackHandler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed.query)
        self.server.auth_code = params.get("code", [None])[0]
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"<h1>You may close this window.</h1>")

def start_server():
    server = http.server.HTTPServer(("localhost", 8000), OAuthCallbackHandler)
    server.handle_request()
    return server.auth_code

# 3. Direct user to log in
auth_url = (
    "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?"
    f"client_id={client_id}&response_type=code&redirect_uri={redirect_uri}"
    f"&response_mode=query&scope={' '.join(scopes)}&state=12345"
)
print("🌐 Opening browser for login...")
webbrowser.open(auth_url)

# 4. Wait for login and get code
print("🔁 Waiting for authorization code...")
auth_code = start_server()
print(f"✅ Code received: {auth_code}")

# 5. Exchange code for token
token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
token_data = {
    "client_id": client_id,
    "client_secret": client_secret,  
    "redirect_uri": redirect_uri,
    "grant_type": "authorization_code",
    "code": auth_code,
    "scope": " ".join(scopes),
}
token_response = requests.post(token_url, data=token_data).json()
access_token = token_response.get("access_token")

if not access_token:
    print("Failed to get access token:", token_response)
    exit()

print("Access token received!")

# 6. Add a task to Microsoft To Do
headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

# First, get your default task list
list_url = "https://graph.microsoft.com/v1.0/me/todo/lists"
response = requests.get(list_url, headers=headers)
if response.status_code != 200:
    print("Failed to get task lists:", response.text)
    exit()

default_list = response.json()["value"][0]["id"]

# Now, add a task
task_url = f"https://graph.microsoft.com/v1.0/me/todo/lists/{default_list}/tasks"
task_data = {
    "title": "Finish AI assistant integration",
    "body": {
        "content": "Don't forget to integrate Microsoft To Do into Joe AI!",
        "contentType": "text"
    }
}

add_response = requests.post(task_url, headers=headers, json=task_data)
if add_response.status_code == 201:
    print("Task added successfully!")
else:
    print("Failed to add task:", add_response.status_code, add_response.text)
