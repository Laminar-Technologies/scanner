import os
import requests
import logging

from asana_error_handler import handle_api_error

BASE_URL = "https://app.asana.com/api/1.0"

class AsanaClient:
    def __init__(self, token, workspace_id):
        self.token = token
        self.workspace_id = workspace_id
        self.base_url = BASE_URL

    def _make_request(self, method, endpoint, params=None, data=None, files=None):
        """Helper to make API requests."""
        url = f"{self.base_url}{endpoint}"
        
        # Use a consistent set of headers, removing Content-Type for file uploads
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Accept": "application/json",
        }
        json_payload = None
        data_payload = None

        if files:
            data_payload = data # For multipart/form-data
        else:
            headers["Content-Type"] = "application/json"
            json_payload = data # For standard JSON posts

        try:
            response = requests.request(method, url, headers=headers, params=params, json=json_payload, data=data_payload, files=files, timeout=30)
            response.raise_for_status()
            
            if response.status_code == 204: # No Content
                return {"success": True, "data": None}
            return {"success": True, "data": response.json()}
        
        except requests.exceptions.RequestException as e:
            return handle_api_error(e, f"{method} {endpoint}")

    # ... (find_user_by_name, find_tag_by_name, find_section_by_name remain the same)

    def find_task_by_wip(self, wip_number, opt_fields="name,gid,parent,memberships"):
        """Searches for a task by its WIP number."""
        logging.info(f"Searching for task with WIP: '{wip_number}'...")
        params = {"text": wip_number, "resource.type": "task", "opt_fields": opt_fields} 
        result = self._make_request('GET', f"/workspaces/{self.workspace_id}/tasks/search", params=params)
        if result["success"] and result["data"]:
            if result["data"].get("data"):
                return {"success": True, "task_data": result["data"]["data"][0]}
            else:
                return {"success": False, "message": f"No task found with WIP: '{wip_number}'."}
        return result

    def get_task_details(self, task_gid, opt_fields="name,gid"):
        """Fetches details for a single task."""
        logging.info(f"Fetching details for task GID: {task_gid}")
        return self._make_request('GET', f"/tasks/{task_gid}", params={"opt_fields": opt_fields})

    def get_subtasks_for_task(self, parent_task_id):
        """Fetches all subtasks for a parent task."""
        logging.info(f"Fetching subtasks for parent task GID: {parent_task_id}")
        return self._make_request('GET', f"/tasks/{parent_task_id}/subtasks", params={"opt_fields": "name,gid"})

    def add_tag_to_task(self, task_id, tag_id):
        """Adds a tag to a task."""
        logging.info(f"Adding tag {tag_id} to task {task_id}")
        return self._make_request('POST', f"/tasks/{task_id}/addTag", data={"data": {"tag": tag_id}})

    def assign_task_to_user(self, task_id, assignee_gid):
        """Assigns a task to a user."""
        logging.info(f"Assigning task {task_id} to user {assignee_gid}")
        return self._make_request('PUT', f"/tasks/{task_id}", data={"data": {"assignee": assignee_gid}})

    def add_comment_to_task(self, task_id, comment_text):
        """Adds a comment to a task. Can be plain text."""
        logging.info(f"Adding comment to task {task_id}")
        # CORRECTED: Use 'text' for plain text comments for simplicity and reliability.
        # Asana will handle basic formatting.
        return self._make_request('POST', f"/tasks/{task_id}/stories", data={"data": {"text": comment_text}})

    def change_task_name(self, task_id, new_name):
        """Changes the name of a task."""
        logging.info(f"Changing name of task {task_id} to '{new_name}'")
        return self._make_request('PUT', f"/tasks/{task_id}", data={"data": {"name": new_name}})

    def move_task_to_section(self, task_id, target_section_id):
        """Moves a task to a section."""
        logging.info(f"Moving task {task_id} to section {target_section_id}")
        return self._make_request('POST', f"/sections/{target_section_id}/addTask", data={"data": {"task": task_id}})
    
    def upload_attachment(self, parent_gid, file_path):
        """Uploads a file as an attachment to a task."""
        logging.info(f"Uploading file '{file_path}' to parent GID: {parent_gid}")
        if not os.path.exists(file_path):
            return {"success": False, "message": f"Attachment file not found at: {file_path}"}
            
        try:
            with open(file_path, 'rb') as f:
                # CORRECTED: The Asana API for attachments does not require a 'data' payload
                # with 'parent' when the endpoint already specifies the task GID.
                # It just needs the file itself.
                files_payload = {'file': (os.path.basename(file_path), f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
                return self._make_request('POST', f"/tasks/{parent_gid}/attachments", files=files_payload)
        except Exception as e:
            return {"success": False, "message": f"Error opening or reading file {file_path}: {e}"}