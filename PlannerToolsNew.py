# -*- coding: utf-8 -*-
import uuid
from typing import Any, Dict
from mcp.server.fastmcp import FastMCP
import requests
import time
import os
from dotenv import load_dotenv
import json
import sys
from datetime import datetime

# Load environment variables for authentication
load_dotenv()

# Enable verbose logging
DEBUG = True

# Set up logging to file
if DEBUG:
    log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 
                                f"planner_debug_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    sys.stderr = open(log_file_path, 'w')
    print(f"Debug logs will be written to: {log_file_path}", file=sys.stdout)

# Object ID key for Planner objects
PlannerObjectIdKey = "__PlannerObjId"

# Initialize FastMCP server
mcp = FastMCP("planner")

# Global authentication context
access_token = None
token_expiry = 0

# Object cache
object_cache = {}

def log_debug(message):
    """Print debug information when DEBUG is enabled"""
    if DEBUG:
        print(f"DEBUG: {message}", file=sys.stderr)
        
def log_step(step_name, message=""):
    """Log execution steps to track the program flow"""
    if DEBUG:
        step_message = f"STEP: {step_name}"
        if message:
            step_message += f" - {message}"
        print(step_message, file=sys.stderr)

def log_api_call(method, url, status=None, response=None):
    """Log API call details for debugging"""
    if DEBUG:
        log_message = f"API CALL: {method} {url}"
        if status:
            log_message += f" (Status: {status})"
        print(log_message, file=sys.stderr)
        
        if response and hasattr(response, 'json'):
            try:
                # Only log the first few items if it's a collection to avoid overwhelming logs
                data = response.json()
                if isinstance(data, dict) and 'value' in data and isinstance(data['value'], list):
                    count = len(data['value'])
                    if count > 0:
                        log_debug(f"Response contains {count} items")
                        if count > 3:
                            # Show a sample of the data
                            log_debug(f"Sample data: {json.dumps(data['value'][:3], indent=2)}")
                            log_debug("...")
                        else:
                            log_debug(f"Data: {json.dumps(data['value'], indent=2)}")
                else:
                    log_debug(f"Response data: {json.dumps(data, indent=2)}")
            except Exception as e:
                log_debug(f"Could not parse response as JSON: {e}")

def add_object(obj):
    """
    Add an object to the object cache and return a GUID
    This allows objects to be passed between tool calls
    """
    guid = str(uuid.uuid4())
    object_cache[guid] = obj
    log_debug(f"Added object to cache with GUID: {guid[:8]}...")
    return guid

def object_exists(guid):
    """Check if an object with the given GUID exists in the cache"""
    exists = guid in object_cache
    log_debug(f"Checking if object {guid[:8]}... exists: {exists}")
    return exists

def get_object(guid):
    """Retrieve an object from the cache using its GUID"""
    if guid not in object_cache:
        log_debug(f"Object {guid[:8]}... not found in cache")
        raise ValueError(f"Value {guid} is not the GUID of a known object. To pass objects, when calling tools, get the object from Planner and send it as {PlannerObjectIdKey}.")
    
    log_debug(f"Retrieved object {guid[:8]}... from cache")
    return object_cache.get(guid)

def remove_object(guid):
    """Remove an object from the cache to free memory"""
    if guid in object_cache:
        log_debug(f"Removed object {guid[:8]}... from cache")
        del object_cache[guid]
    else:
        log_debug(f"Attempted to remove non-existent object {guid[:8]}... from cache")

def ensure_auth():
    """
    Ensure we have a valid authentication token for Microsoft Graph API
    Authentication flow:
    1. Check if we have a valid cached token
    2. If not, check for a direct token from MS_TOKEN_ID environment variable
    3. If no direct token, use client credentials flow with MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT_ID
    """
    global access_token, token_expiry
    
    log_step("AUTH_CHECK", "Checking if authentication token is needed")
    
    # Check if token is valid
    current_time = time.time()
    if access_token and token_expiry > current_time:
        log_debug(f"Using existing token (expires in {int(token_expiry - current_time)} seconds)")
        return access_token
    
    # First check if a direct token is provided via environment variable
    log_step("AUTH_TOKEN_ENV", "Checking for MS_TOKEN_ID in environment")
    direct_token = os.getenv("MS_TOKEN_ID")
    if direct_token:
        log_debug("Using direct token from MS_TOKEN_ID environment variable")
        access_token = direct_token
        # Set a reasonable expiry time (4 hours) since we don't know the actual expiry
        token_expiry = current_time + 14400  # 4 hours in seconds
        log_step("AUTH_COMPLETE", "Using direct token from environment")
        return access_token
    
    # If no direct token, proceed with client credentials flow
    log_step("AUTH_CLIENT_CREDS", "No direct token found, proceeding with client credentials flow")
    
    # Retrieve credentials from environment variables
    client_id = os.getenv("MS_CLIENT_ID")
    client_secret = os.getenv("MS_CLIENT_SECRET")
    tenant_id = os.getenv("MS_TENANT_ID")
    
    log_debug(f"Using tenant ID: {tenant_id}")
    
    if not all([client_id, client_secret, tenant_id]):
        log_step("AUTH_ERROR", "Missing required environment variables")
        raise ValueError("Missing required environment variables for authentication. Please set MS_CLIENT_ID, MS_CLIENT_SECRET, and MS_TENANT_ID or provide MS_TOKEN_ID directly.")
    
    # Get new token using client credentials flow
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    
    try:
        log_step("AUTH_REQUEST", f"Requesting token from {token_url}")
        log_debug(f"Requesting token with client ID: {client_id}")
        
        response = requests.post(token_url, data=token_data)
        if response.status_code != 200:
            log_step("AUTH_FAILED", f"Token request failed with status {response.status_code}")
            log_debug(f"Token request failed with status {response.status_code}: {response.text}")
            raise Exception(f"Authentication failed: {response.text}")
        
        token_info = response.json()
        access_token = token_info["access_token"]
        # Set expiry time with a 5-minute buffer before the actual expiry
        token_expiry = current_time + token_info.get("expires_in", 3600) - 300  
        
        log_step("AUTH_SUCCESS", f"Token will expire in {token_info.get('expires_in', 3600)} seconds")
        log_debug(f"Received token: {access_token[:10]}...")
        return access_token
    except Exception as e:
        log_step("AUTH_ERROR", f"Authentication error: {str(e)}")
        raise

def call_graph_api(endpoint, method="GET", data=None, params=None, additional_headers=None):
    """
    Make a call to Microsoft Graph API
    
    Args:
        endpoint: The Microsoft Graph API endpoint to call (without the base URL)
        method: HTTP method (GET, POST, PATCH, DELETE)
        data: Request body for POST/PATCH requests
        params: URL parameters
        additional_headers: Additional HTTP headers to include
        
    Returns:
        The JSON response from the API
        
    Raises:
        Exception: For API errors or connection issues
    """
    # Start timing the API call for performance tracking
    start_time = time.time()
    log_step("API_START", f"{method} {endpoint}")
    
    # Get authentication token
    try:
        token = ensure_auth()
    except Exception as e:
        log_step("API_AUTH_ERROR", f"Failed to get authentication token: {str(e)}")
        raise
    
    # Prepare headers
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    
    # Add any additional headers
    if additional_headers:
        headers.update(additional_headers)
        log_debug(f"Added additional headers: {list(additional_headers.keys())}")
    
    # Construct the full URL
    url = f"https://graph.microsoft.com/v1.0/{endpoint}"
    log_api_call(method, url)
    
    try:
        # Execute the appropriate HTTP method
        if method == "GET":
            log_debug(f"Making GET request to {url}" + (f" with params: {params}" if params else ""))
            response = requests.get(url, headers=headers, params=params)
        elif method == "POST":
            log_debug(f"Making POST request to {url}" + (f" with data: {json.dumps(data)}" if data else ""))
            response = requests.post(url, headers=headers, json=data)
        elif method == "PATCH":
            log_debug(f"Making PATCH request to {url}" + (f" with data: {json.dumps(data)}" if data else ""))
            response = requests.patch(url, headers=headers, json=data)
        elif method == "DELETE":
            log_debug(f"Making DELETE request to {url}")
            response = requests.delete(url, headers=headers)
        else:
            log_step("API_ERROR", f"Unsupported method: {method}")
            raise ValueError(f"Unsupported method: {method}")
        
        # Print response info for debugging
        duration = time.time() - start_time
        log_step("API_RESPONSE", f"Status {response.status_code} (took {duration:.2f}s)")
        log_api_call(method, url, response.status_code, response)
        
        # Check for HTTP errors
        response.raise_for_status()
        
        # Return JSON response if available
        if response.content:
            result = response.json()
            return result
        return None
    except requests.exceptions.HTTPError as e:
        # Handle HTTP errors with detailed information
        error_msg = f"API Error: {e.response.status_code}"
        log_step("API_HTTP_ERROR", f"Status {e.response.status_code}")
        
        if e.response.content:
            try:
                error_json = e.response.json()
                error_detail = error_json.get("error", {}).get("message", "Unknown error")
                error_code = error_json.get("error", {}).get("code", "Unknown code")
                error_msg = f"{error_msg} - Code: {error_code}, Message: {error_detail}"
                log_debug(f"Error response: {json.dumps(error_json, indent=2)}")
            except:
                error_msg = f"{error_msg} - {e.response.text}"
                log_debug(f"Error response (raw): {e.response.text}")
        
        log_debug(f"Complete error details: {error_msg}")
        raise Exception(error_msg)
    except Exception as e:
        # Handle other errors (network issues, etc.)
        log_step("API_ERROR", f"Unexpected error: {str(e)}")
        log_debug(f"Exception type: {type(e).__name__}")
        raise
    finally:
        # Always log completion of the API call attempt
        duration = time.time() - start_time
        log_step("API_END", f"{method} {endpoint} (took {duration:.2f}s)")

# Tool: 1
@mcp.tool()
async def planner_get_me():
    """Get information about the current authenticated user"""
    log_step("TOOL_START", "planner_get_me")
    try:
        # Since /me is not available with client credentials, return app info instead
        result = {
            "__PlannerObjType": "AppInfo",
            "message": "This endpoint requires delegated authentication. Using client credentials flow instead.",
            "clientId": os.getenv("MS_CLIENT_ID"),
            "tenantId": os.getenv("MS_TENANT_ID")
        }
        log_step("TOOL_SUCCESS", "planner_get_me - Returned application info")
        return result
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_get_me - {str(e)}")
        return {"error": str(e)}

# Tool: Test API connectivity
@mcp.tool()
async def planner_test_connection():
    """Test the API connection by calling a simple endpoint"""
    log_step("TOOL_START", "planner_test_connection")
    try:
        # Test with organization endpoint which should work with application permissions
        org_data = call_graph_api("organization")
        log_step("TOOL_SUCCESS", "planner_test_connection - API connection successful")
        return {
            "status": "success",
            "message": "API connection successful",
            "data": org_data
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_test_connection - {str(e)}")
        return {"error": str(e)}

# Tool: 2 
@mcp.tool()
async def planner_get_plans():
    """Get all plans that the user has access to"""
    log_step("TOOL_START", "planner_get_plans")
    try:
        log_step("PLANS_FETCH", "Attempting direct plans retrieval")
        
        # Try different approaches to get plans
        try:
            # First try the standard endpoint
            log_step("PLANS_DIRECT", "Trying direct planner/plans endpoint")
            plans_data = call_graph_api("planner/plans")
            plans_list = plans_data.get("value", [])
            
            log_step("PLANS_SUCCESS", f"Found {len(plans_list)} plans directly")
            
            return {
                PlannerObjectIdKey: add_object(plans_list),
                "__PlannerObjType": "Plans",
                "count": len(plans_list),
                "plans": [{"id": plan.get("id"), "title": plan.get("title")} for plan in plans_list]
            }
        except Exception as first_error:
            log_step("PLANS_DIRECT_FAILED", f"Direct plans retrieval failed: {str(first_error)}")
            
            # Try with a different API approach - try getting groups first
            try:
                log_step("PLANS_VIA_GROUPS", "Trying alternative approach via groups")
                groups_data = call_graph_api("groups")
                groups = groups_data.get("value", [])
                log_debug(f"Found {len(groups)} groups to check for plans")
                
                all_plans = []
                group_limit = min(5, len(groups))  # Limit to first 5 groups or fewer
                log_step("PLANS_GROUP_SEARCH", f"Searching through {group_limit} groups")
                
                for i, group in enumerate(groups[:group_limit]):
                    try:
                        group_id = group.get("id")
                        group_name = group.get("displayName", "Unknown")
                        log_debug(f"Checking plans for group {i+1}/{group_limit}: {group_name} ({group_id})")
                        
                        group_plans_data = call_graph_api(f"groups/{group_id}/planner/plans")
                        group_plans = group_plans_data.get("value", [])
                        
                        if group_plans:
                            log_debug(f"Found {len(group_plans)} plans in group '{group_name}'")
                            all_plans.extend(group_plans)
                        else:
                            log_debug(f"No plans found in group '{group_name}'")
                    except Exception as group_error:
                        log_debug(f"Error getting plans for group {group_id}: {str(group_error)}")
                        continue
                
                if all_plans:
                    log_step("PLANS_GROUP_SUCCESS", f"Found {len(all_plans)} plans via groups")
                    return {
                        PlannerObjectIdKey: add_object(all_plans),
                        "__PlannerObjType": "Plans",
                        "count": len(all_plans),
                        "plans": [{"id": plan.get("id"), "title": plan.get("title")} for plan in all_plans]
                    }
                else:
                    # If we still don't have plans, provide a diagnostic message
                    log_step("PLANS_NOT_FOUND", "No plans found via any method")
                    return {
                        "error": "Could not find any plans",
                        "message": "Tried multiple approaches but couldn't retrieve plans.",
                        "original_error": str(first_error),
                        "troubleshooting_steps": [
                            "1. Verify API permissions in Azure: Tasks.ReadWrite.All and Group.Read.All",
                            "2. Grant admin consent for these permissions",
                            "3. Check that Microsoft Planner is enabled for your tenant",
                            "4. Verify that you have plans created in Microsoft Planner"
                        ]
                    }
            except Exception as second_error:
                log_step("PLANS_VIA_GROUPS_FAILED", f"Group-based approach failed: {str(second_error)}")
                # If both approaches fail, provide detailed error information
                return {
                    "error": "Failed to retrieve plans",
                    "message": "Both direct and group-based approaches failed to retrieve plans.",
                    "first_attempt_error": str(first_error),
                    "second_attempt_error": str(second_error),
                    "troubleshooting_steps": [
                        "1. Verify API permissions in Azure: Tasks.ReadWrite.All and Group.Read.All",
                        "2. Grant admin consent for these permissions",
                        "3. Check that Microsoft Planner is enabled for your tenant",
                        "4. Verify that you have plans created in Microsoft Planner"
                    ]
                }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_get_plans - Unexpected error: {str(e)}")
        return {"error": str(e)}

# Tool: 3
@mcp.tool()
async def planner_get_plan_details(plan_id: str):
    """Get details for a specific plan"""
    log_step("TOOL_START", f"planner_get_plan_details for plan {plan_id}")
    try:
        log_debug(f"Fetching details for plan ID: {plan_id}")
        plan_data = call_graph_api(f"planner/plans/{plan_id}")
        
        log_step("TOOL_SUCCESS", f"Retrieved plan details for '{plan_data.get('title')}'")
        return {
            PlannerObjectIdKey: add_object(plan_data),
            "__PlannerObjType": "Plan",
            "id": plan_data.get("id"),
            "title": plan_data.get("title"),
            "owner": plan_data.get("owner")
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_get_plan_details - {str(e)}")
        return {"error": str(e)}

# Tool: 4
@mcp.tool()
async def planner_get_tasks(plan_id: str):
    """Get tasks for a specific plan"""
    log_step("TOOL_START", f"planner_get_tasks for plan {plan_id}")
    try:
        log_debug(f"Fetching tasks for plan ID: {plan_id}")
        tasks_data = call_graph_api(f"planner/plans/{plan_id}/tasks")
        tasks_list = tasks_data.get("value", [])
        
        log_step("TOOL_SUCCESS", f"Retrieved {len(tasks_list)} tasks for plan {plan_id}")
        return {
            PlannerObjectIdKey: add_object(tasks_list),
            "__PlannerObjType": "Tasks",
            "count": len(tasks_list),
            "tasks": [{"id": task.get("id"), "title": task.get("title"), "percentComplete": task.get("percentComplete")} 
                    for task in tasks_list]
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_get_tasks - {str(e)}")
        return {"error": str(e)}

# Tool: 5
@mcp.tool()
async def planner_create_task(plan_id: str, title: str, bucket_id: str = None):
    """Create a new task in a plan"""
    log_step("TOOL_START", f"planner_create_task in plan {plan_id}")
    try:
        task_data = {
            "planId": plan_id,
            "title": title
        }
        
        if bucket_id:
            task_data["bucketId"] = bucket_id
        
        new_task = call_graph_api("planner/tasks", method="POST", data=task_data)
        
        log_step("TOOL_SUCCESS", f"Created task '{title}' in plan {plan_id}")
        return {
            PlannerObjectIdKey: add_object(new_task),
            "__PlannerObjType": "Task",
            "id": new_task.get("id"),
            "title": new_task.get("title"),
            "planId": new_task.get("planId")
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_create_task - {str(e)}")
        return {"error": str(e)}

# Tool: 6
@mcp.tool()
async def planner_get_buckets(plan_id: str):
    """Get buckets for a specific plan"""
    log_step("TOOL_START", f"planner_get_buckets for plan {plan_id}")
    try:
        buckets_data = call_graph_api(f"planner/plans/{plan_id}/buckets")
        buckets_list = buckets_data.get("value", [])
        
        log_step("TOOL_SUCCESS", f"Retrieved {len(buckets_list)} buckets for plan {plan_id}")
        return {
            PlannerObjectIdKey: add_object(buckets_list),
            "__PlannerObjType": "Buckets",
            "count": len(buckets_list),
            "buckets": [{"id": bucket.get("id"), "name": bucket.get("name")} for bucket in buckets_list]
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_get_buckets - {str(e)}")
        return {"error": str(e)}

# Tool: 7
@mcp.tool()
async def planner_create_bucket(plan_id: str, name: str):
    """Create a new bucket in a plan"""
    log_step("TOOL_START", f"planner_create_bucket in plan {plan_id}")
    try:
        bucket_data = {
            "planId": plan_id,
            "name": name
        }
        
        new_bucket = call_graph_api("planner/buckets", method="POST", data=bucket_data)
        
        log_step("TOOL_SUCCESS", f"Created bucket '{name}' in plan {plan_id}")
        return {
            PlannerObjectIdKey: add_object(new_bucket),
            "__PlannerObjType": "Bucket",
            "id": new_bucket.get("id"),
            "name": new_bucket.get("name"),
            "planId": new_bucket.get("planId")
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_create_bucket - {str(e)}")
        return {"error": str(e)}

# Tool: 8
@mcp.tool()
async def planner_update_task_progress(task_id: str, percent_complete: int):
    """Update the progress of a task"""
    log_step("TOOL_START", f"planner_update_task_progress for task {task_id}")
    try:
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
        
        log_step("TOOL_SUCCESS", f"Updated task {task_id} progress to {percent_complete}%")
        return {
            PlannerObjectIdKey: add_object(updated_task or task),
            "__PlannerObjType": "Task",
            "id": task_id,
            "title": task.get("title"),
            "percentComplete": percent_complete
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_update_task_progress - {str(e)}")
        return {"error": str(e)}

# Tool: 9
@mcp.tool()
async def planner_create_plan(group_id: str, title: str):
    """Create a new plan in a Microsoft 365 Group"""
    log_step("TOOL_START", f"planner_create_plan in group {group_id}")
    try:
        plan_data = {
            "owner": group_id,
            "title": title
        }
        
        new_plan = call_graph_api("planner/plans", method="POST", data=plan_data)
        
        log_step("TOOL_SUCCESS", f"Created plan '{title}' in group {group_id}")
        return {
            PlannerObjectIdKey: add_object(new_plan),
            "__PlannerObjType": "Plan",
            "id": new_plan.get("id"),
            "title": new_plan.get("title"),
            "owner": new_plan.get("owner")
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_create_plan - {str(e)}")
        return {"error": str(e)}

# Tool: 10
@mcp.tool()
async def planner_get_my_groups():
    """Get all groups that the user is a member of"""
    log_step("TOOL_START", "planner_get_my_groups")
    try:
        # With app permissions, we get all groups instead of just user's groups
        groups_data = call_graph_api("groups")
        groups_list = groups_data.get("value", [])
        
        # Filter to only include Microsoft 365 groups (which can be used with Planner)
        planner_groups = [g for g in groups_list if g.get("groupTypes") and "unified" in g.get("groupTypes")]
        
        log_step("TOOL_SUCCESS", f"Retrieved {len(planner_groups)} groups")
        return {
            PlannerObjectIdKey: add_object(planner_groups),
            "__PlannerObjType": "Groups",
            "count": len(planner_groups),
            "groups": [{"id": group.get("id"), "displayName": group.get("displayName")} for group in planner_groups]
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_get_my_groups - {str(e)}")
        return {"error": str(e)}

# Tool: 11
@mcp.tool()
async def planner_get_task_details(task_id: str):
    """Get detailed information about a specific task"""
    log_step("TOOL_START", f"planner_get_task_details for task {task_id}")
    try:
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
        
        log_step("TOOL_SUCCESS", f"Retrieved details for task {task_id}")
        return {
            PlannerObjectIdKey: add_object(combined_data),
            "__PlannerObjType": "TaskDetails",
            **combined_data
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_get_task_details - {str(e)}")
        return {"error": str(e)}

# Tool: 12
@mcp.tool()
async def planner_set_due_date(task_id: str, due_date_time: str):
    """Set the due date for a task (ISO 8601 format)"""
    log_step("TOOL_START", f"planner_set_due_date for task {task_id}")
    try:
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
        
        log_step("TOOL_SUCCESS", f"Set due date for task {task_id} to {due_date_time}")
        return {
            PlannerObjectIdKey: add_object(updated_task or task),
            "__PlannerObjType": "Task",
            "id": task_id,
            "title": task.get("title"),
            "dueDateTime": due_date_time
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_set_due_date - {str(e)}")
        return {"error": str(e)}

# Tool: 13
@mcp.tool()
async def planner_assign_task(task_id: str, user_id: str):
    """Assign a task to a user"""
    log_step("TOOL_START", f"planner_assign_task for task {task_id}")
    try:
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
        
        log_step("TOOL_SUCCESS", f"Assigned task {task_id} to user {user_id}")
        return {
            PlannerObjectIdKey: add_object(updated_task or task),
            "__PlannerObjType": "Task",
            "id": task_id,
            "title": task.get("title"),
            "assignments": assignments
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_assign_task - {str(e)}")
        return {"error": str(e)}

# Tool: 14
@mcp.tool()
async def planner_delete_task(task_id: str):
    """Delete a task"""
    log_step("TOOL_START", f"planner_delete_task for task {task_id}")
    try:
        # First get the current etag
        task = call_graph_api(f"planner/tasks/{task_id}")
        etag = task.get("@odata.etag")
        
        headers = {
            "If-Match": etag
        }
        
        call_graph_api(f"planner/tasks/{task_id}", method="DELETE", additional_headers=headers)
        
        log_step("TOOL_SUCCESS", f"Deleted task {task_id}")
        return {
            "message": f"Task {task_id} deleted successfully"
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_delete_task - {str(e)}")
        return {"error": str(e)}

# Tool: 15
@mcp.tool()
async def planner_update_task_description(task_id: str, description: str):
    """Update the description of a task"""
    log_step("TOOL_START", f"planner_update_task_description for task {task_id}")
    try:
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
        
        log_step("TOOL_SUCCESS", f"Updated description for task {task_id}")
        return {
            PlannerObjectIdKey: add_object(updated_details or task_details),
            "__PlannerObjType": "TaskDetails",
            "id": task_id,
            "description": description
        }
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_update_task_description - {str(e)}")
        return {"error": str(e)}

# Tool: Check for plans with diagnostics
@mcp.tool()
async def planner_find_plans():
    """Find available plans with detailed diagnostics"""
    log_step("TOOL_START", "planner_find_plans")
    results = {
        "api_tests": {},
        "plans": [],
        "groups": [],
        "advice": []
    }
    
    # Step 1: Check if we can access the organization
    try:
        log_step("ORG_ACCESS_TEST", "Testing organization access")
        org_data = call_graph_api("organization")
        results["api_tests"]["organization"] = "Success"
        results["tenant_id"] = org_data.get("value", [{}])[0].get("id", "Unknown")
    except Exception as e:
        log_step("ORG_ACCESS_FAILED", f"Failed to access organization: {str(e)}")
        results["api_tests"]["organization"] = f"Failed: {str(e)}"
    
    # Step 2: Check if we can access groups
    try:
        log_step("GROUP_ACCESS_TEST", "Testing groups access")
        groups_data = call_graph_api("groups")
        groups = groups_data.get("value", [])
        results["api_tests"]["groups"] = f"Success ({len(groups)} groups found)"
        results["groups"] = [{"id": g.get("id"), "name": g.get("displayName")} for g in groups[:5]]
        
        if not groups:
            log_debug("No Microsoft 365 Groups found in tenant")
            results["advice"].append("No Microsoft 365 Groups found in your tenant. Plans are typically associated with groups.")
            results["advice"].append("Try creating a Microsoft 365 Group from the Microsoft 365 admin center first.")
    except Exception as e:
        log_step("GROUP_ACCESS_FAILED", f"Failed to access groups: {str(e)}")
        results["api_tests"]["groups"] = f"Failed: {str(e)}"
    
    # Step 3: Try direct plan access
    try:
        log_step("PLAN_ACCESS_TEST", "Testing direct plan access")
        plans_data = call_graph_api("planner/plans")
        plans = plans_data.get("value", [])
        results["api_tests"]["plans"] = f"Success ({len(plans)} plans found)"
        results["plans"] = [{"id": p.get("id"), "title": p.get("title")} for p in plans]
        
        if not plans:
            log_debug("No plans found directly")
            results["advice"].append("No plans found. You may need to create plans using the Microsoft Planner web interface.")
    except Exception as e:
        err_msg = str(e)
        log_step("PLAN_ACCESS_FAILED", f"Failed to access plans: {err_msg}")
        results["api_tests"]["plans"] = f"Failed: {err_msg}"
        
        if "Tenant is not found" in err_msg:
            log_debug("Tenant not found error detected")
            results["advice"].append("The 'Tenant is not found' error typically means:")
            results["advice"].append("1. Your tenant doesn't have Planner service enabled")
            results["advice"].append("2. Your application doesn't have Tasks.ReadWrite.All permission")
            results["advice"].append("3. Admin consent hasn't been granted for the required permissions")
    
    # Step 4: Try to check plan service configuration
    if results["groups"]:
        try:
            group_id = results["groups"][0]["id"]
            log_step("GROUP_PLAN_TEST", f"Testing group plans access for group {group_id}")
            group_plans = call_graph_api(f"groups/{group_id}/planner/plans")
            results["api_tests"]["group_plans"] = f"Success ({len(group_plans.get('value', []))} group plans found)"
        except Exception as e:
            log_step("GROUP_PLAN_FAILED", f"Failed to access group plans: {str(e)}")
            results["api_tests"]["group_plans"] = f"Failed: {str(e)}"
    
    # Add general advice
    if not any("Success" in v and "plans found" in v for v in results["api_tests"].values()):
        log_debug("No plans could be accessed through any method")
        results["advice"].append("No plans could be accessed through any method.")
        results["advice"].append("Verify Microsoft Planner is enabled for your tenant.")
        results["advice"].append("Check that your application has Tasks.ReadWrite.All and Group.Read.All permissions.")
        results["advice"].append("Ensure admin consent has been granted for these permissions.")
    
    log_step("TOOL_SUCCESS", "planner_find_plans completed")
    return results

# Tool: Generate status report for a plan
@mcp.tool()
async def planner_generate_status_report(plan_id: str, include_details: bool = False):
    """Generate a comprehensive status report for a plan"""
    log_step("TOOL_START", f"planner_generate_status_report for plan {plan_id}")
    try:
        from datetime import datetime, timezone, timedelta
        
        # Step 1: Get plan details
        log_step("STATUS_PLAN_DETAILS", "Fetching plan details")
        plan_data = call_graph_api(f"planner/plans/{plan_id}")
        
        # Step 2: Get all tasks for the plan
        log_step("STATUS_TASKS", "Fetching tasks")
        tasks_data = call_graph_api(f"planner/plans/{plan_id}/tasks")
        tasks = tasks_data.get("value", [])
        
        # Step 3: Get buckets for better categorization
        log_step("STATUS_BUCKETS", "Fetching buckets")
        buckets_data = call_graph_api(f"planner/plans/{plan_id}/buckets")
        buckets = buckets_data.get("value", [])
        bucket_map = {b["id"]: b["name"] for b in buckets}
        
        # Step 4: Fetch task details if requested
        task_details_map = {}
        if include_details and tasks:
            log_step("STATUS_TASK_DETAILS", "Fetching detailed task information")
            for task in tasks[:min(len(tasks), 10)]:  # Limit to 10 tasks to avoid API throttling
                try:
                    task_id = task.get("id")
                    task_details = call_graph_api(f"planner/tasks/{task_id}/details")
                    task_details_map[task_id] = task_details
                except Exception as e:
                    log_debug(f"Error fetching details for task {task_id}: {str(e)}")
        
        # Step 5: Analyze tasks and organize by status
        now = datetime.now(timezone.utc)
        today = now.replace(hour=0, minute=0, second=0, microsecond=0)
        tomorrow = today + timedelta(days=1)
        next_week = today + timedelta(days=7)
        
        completed_tasks = []
        overdue_tasks = []
        due_today_tasks = []
        due_this_week_tasks = []
        upcoming_tasks = []
        no_due_date_tasks = []
        
        for task in tasks:
            task_id = task.get("id")
            title = task.get("title")
            percent_complete = task.get("percentComplete", 0)
            bucket_id = task.get("bucketId")
            bucket_name = bucket_map.get(bucket_id, "Unknown bucket")
            due_date_str = task.get("dueDateTime")
            
            task_info = {
                "id": task_id,
                "title": title,
                "percentComplete": percent_complete,
                "bucket": bucket_name
            }
            
            # Include additional details if available
            if include_details and task_id in task_details_map:
                details = task_details_map[task_id]
                task_info["description"] = details.get("description", "")
                task_info["checklist"] = {
                    "total": len(details.get("checklist", {})),
                    "completed": sum(1 for item in details.get("checklist", {}).values() 
                                   if isinstance(item, dict) and item.get("isChecked", False))
                }
            
            # Check if task is completed
            if percent_complete == 100:
                completed_tasks.append(task_info)
                continue
                
            # Categorize by due date
            if not due_date_str:
                no_due_date_tasks.append(task_info)
                continue
                
            try:
                due_date = datetime.fromisoformat(due_date_str.replace('Z', '+00:00'))
                task_info["dueDate"] = due_date.strftime("%Y-%m-%d")
                
                if due_date < today:
                    overdue_tasks.append(task_info)
                elif due_date < tomorrow:
                    due_today_tasks.append(task_info)
                elif due_date < next_week:
                    due_this_week_tasks.append(task_info)
                else:
                    upcoming_tasks.append(task_info)
            except Exception as e:
                log_debug(f"Error parsing due date {due_date_str} for task {task_id}: {str(e)}")
                no_due_date_tasks.append(task_info)
        
        # Step 6: Calculate summary statistics
        total_tasks = len(tasks)
        completed_count = len(completed_tasks)
        overdue_count = len(overdue_tasks)
        due_today_count = len(due_today_tasks)
        due_this_week_count = len(due_this_week_tasks)
        
        # Overall completion percentage
        completion_percentage = 0
        if total_tasks > 0:
            completion_percentage = int(sum(task.get("percentComplete", 0) for task in tasks) / total_tasks)
        
        # Step 7: Compile the report
        report = {
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "plan": {
                "id": plan_id,
                "title": plan_data.get("title"),
                "owner": plan_data.get("owner"),
                "created_datetime": plan_data.get("createdDateTime")
            },
            "summary": {
                "total_tasks": total_tasks,
                "completed_tasks": completed_count,
                "overdue_tasks": overdue_count,
                "due_today": due_today_count,
                "due_this_week": due_this_week_count,
                "overall_completion_percentage": completion_percentage,
                "buckets": len(buckets)
            },
            "tasks_by_status": {
                "completed": completed_tasks[:10],  # Limit to 10 for readability
                "overdue": overdue_tasks,
                "due_today": due_today_tasks,
                "due_this_week": due_this_week_tasks,
                "upcoming": upcoming_tasks[:10],  # Limit to 10 for readability
                "no_due_date": no_due_date_tasks[:10]  # Limit to 10 for readability
            }
        }
        
        # Step 8: Format the report for better readability
        formatted_report = {
            PlannerObjectIdKey: add_object(report),
            "__PlannerObjType": "StatusReport",
            "plan_title": plan_data.get("title"),
            "generated_at": report["generated_at"],
            "summary": report["summary"],
            "completion_status": f"{report['summary']['completed_tasks']}/{report['summary']['total_tasks']} tasks completed ({report['summary']['overall_completion_percentage']}% overall)",
            "urgent_attention": f"{report['summary']['overdue_tasks']} overdue, {report['summary']['due_today']} due today",
            "tasks_by_status": report["tasks_by_status"],
        }
        
        log_step("TOOL_SUCCESS", f"Status report generated for plan {plan_id}")
        return formatted_report
        
    except Exception as e:
        log_step("TOOL_ERROR", f"planner_generate_status_report - {str(e)}")
        return {"error": str(e)}

if __name__ == "__main__":
    log_step("SERVER_START", "Starting Planner MCP server")
    mcp.run(transport='stdio')
