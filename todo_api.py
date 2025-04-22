import requests

# Retrieves the ID of the default task list for the signed-in Microsoft account
def get_task_list_id(access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get("https://graph.microsoft.com/v1.0/me/todo/lists", headers=headers)
    # Return the ID of the first list if the request is successful
    return response.json()["value"][0]["id"] if response.ok else None

# Adds a new task to the user's default Microsoft To Do list
def add_task(access_token, title, content=""):
    list_id = get_task_list_id(access_token)
    url = f"https://graph.microsoft.com/v1.0/me/todo/lists/{list_id}/tasks"

    # Define task structure including optional body content
    task_data = {
        "title": title,
        "body": {
            "content": content,
            "contentType": "text"
        }
    }
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    return requests.post(url, headers=headers, json=task_data).ok

# Fetches all task titles from the user's default Microsoft To Do list
def get_tasks(access_token):
    list_id = get_task_list_id(access_token)
    url = f"https://graph.microsoft.com/v1.0/me/todo/lists/{list_id}/tasks"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)

    # Return a list of task titles if request is successful
    return [task["title"] for task in response.json()["value"]] if response.ok else []

# Deletes a task from the user's default list by matching part of the task name
def delete_task(access_token, task_name):
    list_id = get_task_list_id(access_token)
    url = f"https://graph.microsoft.com/v1.0/me/todo/lists/{list_id}/tasks"
    headers = {"Authorization": f"Bearer {access_token}"}
    
    # Retrieve all tasks and search for a match with the given name
    tasks = requests.get(url, headers=headers).json()["value"]
    for task in tasks:
        if task_name.lower() in task["title"].lower():
            # Delete the matched task using its ID
            del_url = f"{url}/{task['id']}"
            return requests.delete(del_url, headers=headers).ok

    # Return False if no matching task was found
    return False
