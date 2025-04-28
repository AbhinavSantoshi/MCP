# -*- coding: utf-8 -*-
import uuid
from typing import Any, Dict
from mcp.server.fastmcp import FastMCP
import requests
import time
import os
import json
import webbrowser
import subprocess
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables for authentication
load_dotenv()

# Object ID key for Planner objects
PlannerObjectIdKey = "__PlannerObjId"

# Get preferred browser from environment variables if set
preferred_browser = os.getenv("PLANNER_PREFERRED_BROWSER", "chrome")  # Default to chrome

# Initialize FastMCP server
mcp = FastMCP("planner")

# Global authentication context
access_token = None
refresh_token = None
token_expiry = 0
auth_status = {"initialized": False, "authenticated": False}
token_cache_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.token_cache')

# Object cache
object_cache = {}

# User session info
session_user = None
auth_info = {}

def add_object(obj):
    guid = str(uuid.uuid4())
    object_cache[guid] = obj
    return guid

def object_exists(guid):
    return guid in object_cache

def get_object(guid):
    if guid not in object_cache:
        raise ValueError(f"Value {guid} is not the GUID of a known object. To pass objects, when calling tools, get the object from Planner and send it as {PlannerObjectIdKey}.")
    return object_cache.get(guid)

def remove_object(guid):
    if guid in object_cache:
        del object_cache[guid]

def save_token_cache(token_data):
    """Save token data to cache file"""
    try:
        with open(token_cache_file, 'w') as f:
            json.dump(token_data, f)
    except Exception as e:
        print(f"Warning: Could not save token cache: {e}")

def load_token_cache():
    """Load token data from cache file"""
    try:
        if os.path.exists(token_cache_file):
            with open(token_cache_file, 'r') as f:
                return json.load(f)
    except Exception as e:
        print(f"Warning: Could not load token cache: {e}")
    return None

def check_auth_status():
    """Check if user is authenticated and return status"""
    global auth_status, access_token, token_expiry
    
    current_time = time.time()
    auth_status["authenticated"] = access_token is not None and token_expiry > current_time
    return auth_status

def initialize_session():
    """Initialize a new user session, loading cached credentials if available"""
    global session_user, auth_status
    
    # Try to load tokens from cache
    token_cache = load_token_cache()
    if token_cache:
        # Get user info if we have valid tokens
        try:
            ensure_auth()  # This will handle token refresh if needed
            user_data = call_graph_api("me")
            session_user = {
                "id": user_data.get("id"),
                "displayName": user_data.get("displayName"),
                "userPrincipalName": user_data.get("userPrincipalName"),
                "sessionStartTime": time.time()
            }
            auth_status["initialized"] = True
            return True
        except Exception as e:
            print(f"Warning: Could not initialize session from cached tokens: {e}")
    
    auth_status["initialized"] = False
    return False

def logout():
    """Clear current session and remove cached tokens"""
    global access_token, refresh_token, token_expiry, session_user, auth_status
    
    # Clear in-memory data
    access_token = None
    refresh_token = None
    token_expiry = 0
    session_user = None
    auth_status = {"initialized": False, "authenticated": False}
    
    # Remove token cache file
    try:
        if os.path.exists(token_cache_file):
            os.remove(token_cache_file)
            return {"message": "Successfully logged out and cleared cached credentials"}
    except Exception as e:
        return {"message": f"Error during logout: {e}"}
    
    return {"message": "Logged out"}

def ensure_auth(browser=None):
    """Ensure we have a valid authentication token for Microsoft Graph API
    
    Args:
        browser (str, optional): Browser to use for authentication. 
            Options include: 'chrome', 'firefox', 'edge', 'safari', etc.
            If not specified, uses the browser from PLANNER_PREFERRED_BROWSER env var
            or system default if not set.
    """
    global access_token, refresh_token, token_expiry, preferred_browser, auth_status, session_user, auth_info
    
    # Use provided browser or environment variable
    browser_to_use = browser or preferred_browser
    
    # Check if token is valid
    current_time = time.time()
    if access_token and token_expiry > current_time:
        auth_status["authenticated"] = True
        print("Reusing existing credentials.")
        return access_token
    
    # Try to load from cache first
    if not access_token and not refresh_token:
        token_cache = load_token_cache()
        if token_cache:
            access_token = token_cache.get("access_token")
            refresh_token = token_cache.get("refresh_token")
            token_expiry = token_cache.get("expiry", 0)
            
            # Load cached client_id and tenant_id if available
            if "client_id" in token_cache and "tenant_id" in token_cache:
                auth_info["client_id"] = token_cache.get("client_id")
                auth_info["tenant_id"] = token_cache.get("tenant_id")
            
            # If token is still valid, return it
            if access_token and token_expiry > current_time:
                auth_status["authenticated"] = True
                print("Reusing cached credentials.")
                return access_token
    
    # Try to use refresh token if available
    if refresh_token:
        try:
            client_id = os.getenv("MS_CLIENT_ID")
            tenant_id = os.getenv("MS_TENANT_ID")
            
            token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
            token_data = {
                'grant_type': 'refresh_token',
                'client_id': client_id,
                'refresh_token': refresh_token,
                'scope': 'User.Read Group.Read.All Tasks.ReadWrite'
            }
            
            response = requests.post(token_url, data=token_data)
            if response.status_code == 200:
                token_info = response.json()
                access_token = token_info["access_token"]
                refresh_token = token_info.get("refresh_token", refresh_token)
                token_expiry = current_time + token_info.get("expires_in", 3600) - 300  # Buffer of 5 minutes
                
                # Save to cache
                save_token_cache({
                    "access_token": access_token,
                    "refresh_token": refresh_token,
                    "expiry": token_expiry
                })
                
                auth_status["authenticated"] = True
                print("Successfully refreshed credentials.")
                return access_token
        except Exception as e:
            print(f"Error refreshing token: {e}")
            # Clear tokens if refresh fails
            access_token = None
            refresh_token = None
    
    # If no valid token yet, guide user through interactive authentication
    client_id = os.getenv("MS_CLIENT_ID")
    tenant_id = os.getenv("MS_TENANT_ID")
    # Using a commonly accepted redirect URI for native applications
    redirect_uri = "http://localhost"
    
    # Build the authorization URL
    auth_url = (
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/authorize"
        f"?client_id={client_id}"
        f"&response_type=code"
        f"&redirect_uri={redirect_uri}"
        f"&response_mode=query"
        f"&scope=User.Read%20Group.Read.All%20Tasks.ReadWrite"
    )
    
    print("\n" + "="*80)
    print("To use the Planner API, please authenticate with your Microsoft account.")
    print(f"Authentication URL: {auth_url}")
    
    # Find available browsers
    available_browsers = {}
    chrome_names = ["chrome", "chrome-browser", "google-chrome", "googlechrome", "Google Chrome"]
    
    for browser_name in webbrowser._browsers:
        if browser_name:  # Skip empty names
            available_browsers[browser_name.lower()] = browser_name
    
    # Special handling for Chrome
    detected_chrome = None
    for chrome_name in chrome_names:
        if chrome_name.lower() in available_browsers:
            detected_chrome = available_browsers[chrome_name.lower()]
            break
    
    # Try to find Chrome executable directly on Windows
    if not detected_chrome and os.name == 'nt':
        chrome_paths = [
            os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe")
        ]
        
        for path in chrome_paths:
            if os.path.exists(path):
                try:
                    # Register Chrome browser manually
                    chrome_path = f'"{path}" %s'
                    webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
                    detected_chrome = 'chrome'
                    available_browsers['chrome'] = 'chrome'
                    print(f"Registered Chrome at: {path}")
                    break
                except Exception as e:
                    print(f"Failed to register Chrome at {path}: {e}")
    
    # Print browser info
    print("\nAvailable browsers:")
    for name in available_browsers.values():
        print(f"- {name}")
    
    if browser_to_use:
        print(f"\nAttempting to use {browser_to_use} browser...")
    else:
        print("\nNo specific browser requested. Will try Chrome first, then system default.")
        
    # Try to open the browser in this order:
    # 1. Requested browser (if specified)
    # 2. Chrome (if detected)
    # 3. System default
    browser_opened = False
    
    try:
        if browser_to_use:
            # Try to use the requested browser
            try:
                if browser_to_use.lower() == 'chrome' and os.name == 'nt':
                    # Direct launch of Chrome with subprocess on Windows
                    chrome_paths = [
                        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
                        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
                        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe")
                    ]
                    
                    for path in chrome_paths:
                        if os.path.exists(path):
                            try:
                                subprocess.Popen([path, auth_url])
                                print(f"Opened Chrome browser directly using subprocess.")
                                browser_opened = True
                                break
                            except Exception as e:
                                print(f"Failed to open Chrome directly: {e}")
                
                # If direct Chrome launch failed, try through webbrowser module
                if not browser_opened:
                    browser_controller = webbrowser.get(browser_to_use)
                    browser_opened = browser_controller.open(auth_url)
                    print(f"Opened {browser_to_use} browser using webbrowser module.")
            except Exception as e:
                print(f"Could not open {browser_to_use} browser: {e}")
        
        # If the requested browser failed or none was specified, try Chrome
        if not browser_opened and detected_chrome:
            try:
                if os.name == 'nt':
                    # Try direct Chrome launch again
                    for path in chrome_paths:
                        if os.path.exists(path):
                            try:
                                subprocess.Popen([path, auth_url])
                                print(f"Opened Chrome browser directly using subprocess.")
                                browser_opened = True
                                break
                            except Exception as e:
                                print(f"Failed to open Chrome directly: {e}")
                
                # If direct launch failed, try through webbrowser module
                if not browser_opened:
                    browser_controller = webbrowser.get(detected_chrome)
                    browser_opened = browser_controller.open(auth_url)
                    print(f"Opened Chrome browser using webbrowser module.")
            except Exception as e:
                print(f"Could not open Chrome: {e}")
        
        # If all else fails, try the system default
        if not browser_opened:
            browser_opened = webbrowser.open(auth_url)
            print("Opened default system browser.")
            
    except Exception as e:
        print(f"Could not open any browser automatically: {e}")
    
    if not browser_opened:
        print("\nUnable to open browser automatically. Please copy and open this URL manually:")
        print(auth_url)
    
    print("\nAfter signing in, you will be redirected to a page with a URL like:")
    print("https://login.microsoftonline.com/common/oauth2/nativeclient?code=SOME_CODE")
    print("Please copy the ENTIRE URL from your browser and paste it below.")
    print("="*80 + "\n")
    
    # Get the authorization code from the user
    auth_response_url = input("Paste the full URL after authentication: ")
    
    # Extract the authorization code from the URL
    try:
        if "code=" in auth_response_url:
            auth_code = auth_response_url.split("code=")[1].split("&")[0]
        else:
            raise ValueError("No authorization code found in the URL")
        
        # Exchange the code for tokens
        token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        token_data = {
            'grant_type': 'authorization_code',
            'client_id': client_id,
            'code': auth_code,
            'redirect_uri': redirect_uri,
            'scope': 'User.Read Group.Read.All Tasks.ReadWrite'
        }
        
        response = requests.post(token_url, data=token_data)
        response.raise_for_status()
        
        token_info = response.json()
        access_token = token_info["access_token"]
        refresh_token = token_info.get("refresh_token")
        token_expiry = current_time + token_info.get("expires_in", 3600) - 300  # Buffer of 5 minutes
        
        # Extract auth info from response
        auth_info["client_id"] = client_id
        auth_info["tenant_id"] = tenant_id
        
        # Try to extract client ID and tenant ID from the token if possible
        try:
            import jwt
            # Decode JWT token without verification to extract claims
            decoded_token = jwt.decode(access_token, options={"verify_signature": False})
            if "tid" in decoded_token:
                auth_info["tenant_id"] = decoded_token["tid"]
                print(f"Extracted tenant ID from token: {auth_info['tenant_id']}")
            if "appid" in decoded_token:
                auth_info["client_id"] = decoded_token["appid"]
                print(f"Extracted client ID from token: {auth_info['client_id']}")
            if "aud" in decoded_token and auth_info.get("client_id") is None:
                # Sometimes the audience field contains the client ID
                auth_info["client_id"] = decoded_token["aud"]
                print(f"Extracted client ID from token audience: {auth_info['client_id']}")
        except ImportError:
            print("JWT module not found, using provided client ID and tenant ID")
        except Exception as e:
            print(f"Could not extract auth info from token: {e}")
        
        # Save to cache
        save_token_cache({
            "access_token": access_token,
            "refresh_token": refresh_token,
            "expiry": token_expiry,
            "client_id": auth_info["client_id"],
            "tenant_id": auth_info["tenant_id"]
        })
        
        auth_status["authenticated"] = True
        print("Authentication successful!")
        return access_token
        
    except Exception as e:
        print(f"Error during authentication: {str(e)}")
        raise Exception("Authentication failed. Please try again.")

def call_graph_api(endpoint, method="GET", data=None, params=None, additional_headers=None):
    """Make a call to Microsoft Graph API"""
    token = ensure_auth()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    # Add any additional headers
    if additional_headers:
        headers.update(additional_headers)
    
    url = f"https://graph.microsoft.com/v1.0/{endpoint}"
    
    try:
        if method == "GET":
            response = requests.get(url, headers=headers, params=params)
        elif method == "POST":
            response = requests.post(url, headers=headers, json=data)
        elif method == "PATCH":
            response = requests.patch(url, headers=headers, json=data)
        elif method == "DELETE":
            response = requests.delete(url, headers=headers)
        else:
            raise ValueError(f"Unsupported method: {method}")
        
        # Check for HTTP errors
        response.raise_for_status()
        
        if response.content:
            return response.json()
        return None
    except requests.exceptions.HTTPError as e:
        error_msg = f"API Error: {e.response.status_code}"
        if e.response.content:
            try:
                error_json = e.response.json()
                error_detail = error_json.get("error", {}).get("message", "Unknown error")
                error_msg = f"{error_msg} - {error_detail}"
            except:
                error_msg = f"{error_msg} - {e.response.text}"
        raise Exception(error_msg)

# Tool: 1
@mcp.tool()
async def planner_get_me():
    """Get information about the current authenticated user"""
    user_data = call_graph_api("me")
    return {
        PlannerObjectIdKey: add_object(user_data),
        "__PlannerObjType": "User",
        "displayName": user_data.get("displayName"),
        "userPrincipalName": user_data.get("userPrincipalName"),
        "id": user_data.get("id")
    }

# Tool: 2 
@mcp.tool()
async def planner_get_plans():
    """Get all plans that the user has access to"""
    plans_data = call_graph_api("me/planner/plans")
    plans_list = plans_data.get("value", [])
    
    return {
        PlannerObjectIdKey: add_object(plans_list),
        "__PlannerObjType": "Plans",
        "count": len(plans_list),
        "plans": [{"id": plan.get("id"), "title": plan.get("title")} for plan in plans_list]
    }

# Tool: 3
@mcp.tool()
async def planner_get_plan_details(plan_id: str):
    """Get details for a specific plan"""
    plan_data = call_graph_api(f"planner/plans/{plan_id}")
    
    return {
        PlannerObjectIdKey: add_object(plan_data),
        "__PlannerObjType": "Plan",
        "id": plan_data.get("id"),
        "title": plan_data.get("title"),
        "owner": plan_data.get("owner")
    }

# Tool: 4
@mcp.tool()
async def planner_get_tasks(plan_id: str):
    """Get tasks for a specific plan"""
    tasks_data = call_graph_api(f"planner/plans/{plan_id}/tasks")
    tasks_list = tasks_data.get("value", [])
    
    return {
        PlannerObjectIdKey: add_object(tasks_list),
        "__PlannerObjType": "Tasks",
        "count": len(tasks_list),
        "tasks": [{"id": task.get("id"), "title": task.get("title"), "percentComplete": task.get("percentComplete")} 
                 for task in tasks_list]
    }

# Tool: 5
@mcp.tool()
async def planner_create_task(plan_id: str, title: str, bucket_id: str = None):
    """Create a new task in a plan"""
    task_data = {
        "planId": plan_id,
        "title": title
    }
    
    if bucket_id:
        task_data["bucketId"] = bucket_id
    
    new_task = call_graph_api("planner/tasks", method="POST", data=task_data)
    
    return {
        PlannerObjectIdKey: add_object(new_task),
        "__PlannerObjType": "Task",
        "id": new_task.get("id"),
        "title": new_task.get("title"),
        "planId": new_task.get("planId")
    }

# Tool: 6
@mcp.tool()
async def planner_get_buckets(plan_id: str):
    """Get buckets for a specific plan"""
    buckets_data = call_graph_api(f"planner/plans/{plan_id}/buckets")
    buckets_list = buckets_data.get("value", [])
    
    return {
        PlannerObjectIdKey: add_object(buckets_list),
        "__PlannerObjType": "Buckets",
        "count": len(buckets_list),
        "buckets": [{"id": bucket.get("id"), "name": bucket.get("name")} for bucket in buckets_list]
    }

# Tool: 7
@mcp.tool()
async def planner_create_bucket(plan_id: str, name: str):
    """Create a new bucket in a plan"""
    bucket_data = {
        "planId": plan_id,
        "name": name
    }
    
    new_bucket = call_graph_api("planner/buckets", method="POST", data=bucket_data)
    
    return {
        PlannerObjectIdKey: add_object(new_bucket),
        "__PlannerObjType": "Bucket",
        "id": new_bucket.get("id"),
        "name": new_bucket.get("name"),
        "planId": new_bucket.get("planId")
    }

# Tool: 8
@mcp.tool()
async def planner_update_task_progress(task_id: str, percent_complete: int):
    """Update the progress of a task"""
    # First get the current etag
    task = call_graph_api(f"planner/tasks/{task_id}")
    etag = task.get("@odata.etag")
    
    update_data = {
        "percentComplete": percent_complete
    }
    
    # Add the etag as a header for concurrency control
    headers = {
        "If-Match": etag
    }
    
    updated_task = call_graph_api(f"planner/tasks/{task_id}", method="PATCH", 
                                 data=update_data, additional_headers=headers)
    
    return {
        PlannerObjectIdKey: add_object(updated_task or task),
        "__PlannerObjType": "Task",
        "id": task_id,
        "title": task.get("title"),
        "percentComplete": percent_complete
    }

# Tool: 9
@mcp.tool()
async def planner_create_plan(group_id: str, title: str):
    """Create a new plan in a Microsoft 365 Group"""
    plan_data = {
        "owner": group_id,
        "title": title
    }
    
    new_plan = call_graph_api("planner/plans", method="POST", data=plan_data)
    
    return {
        PlannerObjectIdKey: add_object(new_plan),
        "__PlannerObjType": "Plan",
        "id": new_plan.get("id"),
        "title": new_plan.get("title"),
        "owner": new_plan.get("owner")
    }

# Tool: 10
@mcp.tool()
async def planner_get_my_groups():
    """Get all groups that the user is a member of"""
    groups_data = call_graph_api("me/memberOf/$/microsoft.graph.group")
    groups_list = groups_data.get("value", [])
    
    # Filter to only include Microsoft 365 groups (which can be used with Planner)
    planner_groups = [g for g in groups_list if g.get("groupTypes") and "unified" in g.get("groupTypes")]
    
    return {
        PlannerObjectIdKey: add_object(planner_groups),
        "__PlannerObjType": "Groups",
        "count": len(planner_groups),
        "groups": [{"id": group.get("id"), "displayName": group.get("displayName")} for group in planner_groups]
    }

# Tool: 11
@mcp.tool()
async def planner_get_task_details(task_id: str):
    """Get detailed information about a specific task"""
    task = call_graph_api(f"planner/tasks/{task_id}")
    task_details = call_graph_api(f"planner/tasks/{task_id}/details")
    
    # Combine basic task info with details
    combined_data = {
        "id": task.get("id"),
        "title": task.get("title"),
        "planId": task.get("planId"),
        "bucketId": task.get("bucketId"),
        "percentComplete": task.get("percentComplete"),
        "dueDateTime": task.get("dueDateTime"),
        "assignees": task.get("assignments", {}),
        "description": task_details.get("description", ""),
        "checklist": task_details.get("checklist", {}),
        "references": task_details.get("references", {})
    }
    
    return {
        PlannerObjectIdKey: add_object(combined_data),
        "__PlannerObjType": "TaskDetails",
        **combined_data
    }

# Tool: 12
@mcp.tool()
async def planner_set_due_date(task_id: str, due_date_time: str):
    """Set the due date for a task (ISO 8601 format)"""
    # First get the current etag
    task = call_graph_api(f"planner/tasks/{task_id}")
    etag = task.get("@odata.etag")
    
    update_data = {
        "dueDateTime": due_date_time
    }
    
    headers = {
        "If-Match": etag
    }
    
    updated_task = call_graph_api(f"planner/tasks/{task_id}", method="PATCH", 
                                 data=update_data, additional_headers=headers)
    
    return {
        PlannerObjectIdKey: add_object(updated_task or task),
        "__PlannerObjType": "Task",
        "id": task_id,
        "title": task.get("title"),
        "dueDateTime": due_date_time
    }

# Tool: 13
@mcp.tool()
async def planner_assign_task(task_id: str, user_id: str):
    """Assign a task to a user"""
    # First get the current etag
    task = call_graph_api(f"planner/tasks/{task_id}")
    etag = task.get("@odata.etag")
    
    # Create assignment payload
    assignments = task.get("assignments", {})
    assignments[user_id] = {"@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": " !"}
    
    update_data = {
        "assignments": assignments
    }
    
    headers = {
        "If-Match": etag
    }
    
    updated_task = call_graph_api(f"planner/tasks/{task_id}", method="PATCH", 
                                 data=update_data, additional_headers=headers)
    
    return {
        PlannerObjectIdKey: add_object(updated_task or task),
        "__PlannerObjType": "Task",
        "id": task_id,
        "title": task.get("title"),
        "assignments": assignments
    }

# Tool: 14
@mcp.tool()
async def planner_delete_task(task_id: str):
    """Delete a task"""
    # First get the current etag
    task = call_graph_api(f"planner/tasks/{task_id}")
    etag = task.get("@odata.etag")
    
    headers = {
        "If-Match": etag
    }
    
    call_graph_api(f"planner/tasks/{task_id}", method="DELETE", additional_headers=headers)
    
    return {
        "message": f"Task {task_id} deleted successfully"
    }

# Tool: 15
@mcp.tool()
async def planner_update_task_description(task_id: str, description: str):
    """Update the description of a task"""
    # First get the current etag of task details
    task_details = call_graph_api(f"planner/tasks/{task_id}/details")
    etag = task_details.get("@odata.etag")
    
    update_data = {
        "description": description
    }
    
    headers = {
        "If-Match": etag
    }
    
    updated_details = call_graph_api(f"planner/tasks/{task_id}/details", method="PATCH", 
                                   data=update_data, additional_headers=headers)
    
    return {
        PlannerObjectIdKey: add_object(updated_details or task_details),
        "__PlannerObjType": "TaskDetails",
        "id": task_id,
        "description": description
    }

# Tool: 16 (new tool for login)
@mcp.tool()
async def planner_login(force_new_login: bool = False):
    """Log in to Microsoft Planner. 
    Will use cached credentials if available unless force_new_login is set to True."""
    global session_user, auth_status
    
    if force_new_login:
        # Clear existing tokens to force a new login
        logout()
    
    # Check if already logged in with valid tokens
    if not force_new_login and check_auth_status()["authenticated"]:
        if session_user:
            return {
                "status": "already_authenticated",
                "message": f"Already logged in as {session_user['displayName']} ({session_user['userPrincipalName']})",
                "user": {
                    "displayName": session_user["displayName"],
                    "userPrincipalName": session_user["userPrincipalName"],
                    "id": session_user["id"]
                }
            }
    
    # Perform authentication and get user info
    try:
        ensure_auth()
        user_data = call_graph_api("me")
        
        # Store user session info
        session_user = {
            "id": user_data.get("id"),
            "displayName": user_data.get("displayName"),
            "userPrincipalName": user_data.get("userPrincipalName"),
            "sessionStartTime": time.time()
        }
        
        auth_status["initialized"] = True
        
        return {
            "status": "success",
            "message": f"Successfully logged in as {user_data.get('displayName')} ({user_data.get('userPrincipalName')})",
            "user": {
                "displayName": user_data.get("displayName"),
                "userPrincipalName": user_data.get("userPrincipalName"),
                "id": user_data.get("id")
            }
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"Login failed: {str(e)}"
        }

# Tool: 17 (new tool for logout)
@mcp.tool()
async def planner_logout():
    """Log out and clear cached credentials"""
    return logout()

if __name__ == "__main__":
    # Try to initialize session from cached credentials on startup
    initialize_session()
    
    if auth_status["initialized"]:
        print(f"Session initialized from cached credentials for {session_user['displayName']}")
        print("Credential reuse is enabled. Call planner_logout() to clear cached credentials.")
    else:
        print("No cached credentials found. You will need to log in when making the first API call.")
    
    mcp.run(transport='stdio')
