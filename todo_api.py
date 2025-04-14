import requests

def get_task_list_id(access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get("https://graph.microsoft.com/v1.0/me/todo/lists", headers=headers)
    return response.json()["value"][0]["id"] if response.ok else None

def add_task(access_token, title, content=""):
    list_id = get_task_list_id(access_token)
    url = f"https://graph.microsoft.com/v1.0/me/todo/lists/{list_id}/tasks"
    task_data = {
        "title": title,
        "body": {
            "content": content,
            "contentType": "text"
        }
    }
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    return requests.post(url, headers=headers, json=task_data).ok

def get_tasks(access_token):
    list_id = get_task_list_id(access_token)
    url = f"https://graph.microsoft.com/v1.0/me/todo/lists/{list_id}/tasks"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    return [task["title"] for task in response.json()["value"]] if response.ok else []

def delete_task(access_token, task_name):
    list_id = get_task_list_id(access_token)
    url = f"https://graph.microsoft.com/v1.0/me/todo/lists/{list_id}/tasks"
    headers = {"Authorization": f"Bearer {access_token}"}
    tasks = requests.get(url, headers=headers).json()["value"]
    for task in tasks:
        if task_name.lower() in task["title"].lower():
            del_url = f"{url}/{task['id']}"
            return requests.delete(del_url, headers=headers).ok
    return False
