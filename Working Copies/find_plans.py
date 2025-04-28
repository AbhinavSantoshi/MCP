#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Find Microsoft Planner plans script - uses the functions from PlannerToolsNew.py
With option to use direct access token
"""

import os
import sys
import argparse
import requests
from dotenv import load_dotenv

# Add the current directory to the path so we can import from PlannerToolsNew
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import functions from the existing file
from PlannerToolsNew import call_graph_api, ensure_auth, log_debug

def call_graph_api_with_token(endpoint, access_token, method="GET", data=None, params=None, additional_headers=None):
    """Make a call to Microsoft Graph API using a provided access token"""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    # Add any additional headers
    if additional_headers:
        headers.update(additional_headers)
    
    url = f"https://graph.microsoft.com/v1.0/{endpoint}"
    print(f"Calling API: {method} {url}")
    
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
        
        print(f"API Response: Status {response.status_code}")
        
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
                error_code = error_json.get("error", {}).get("code", "Unknown code")
                error_msg = f"{error_msg} - Code: {error_code}, Message: {error_detail}"
                print(f"Detailed error: {error_msg}")
            except:
                error_msg = f"{error_msg} - {e.response.text}"
        print(f"Error details: {error_msg}")
        raise Exception(error_msg)
    except Exception as e:
        print(f"API error: {str(e)}")
        raise

def find_plans_with_token(access_token):
    """Find plans using a directly provided access token"""
    print("Using provided access token for authentication")
    
    api_caller = lambda endpoint, method="GET", data=None, params=None, additional_headers=None: call_graph_api_with_token(
        endpoint, access_token, method, data, params, additional_headers
    )
    
    try:
        # Test connection
        print("Checking API connection...")
        org_data = api_caller("organization")
        print(f"✓ API connection successful to tenant: {org_data['value'][0]['displayName']}")
    except Exception as e:
        print(f"✗ API connection failed: {str(e)}")
        return
    
    # Reuse the same pattern as the main function but with our own API caller
    find_plans_implementation(api_caller)

def find_plans_implementation(api_func):
    """Implementation of the plan finding logic, abstracted to work with different API callers"""
    print("\nChecking for Microsoft 365 Groups...")
    try:
        # Get groups
        groups_data = api_func("groups")
        groups = groups_data.get("value", [])
        
        if not groups:
            print("✗ No Microsoft 365 Groups found.")
            print("  Plans are associated with groups. You may need to create a group first.")
        else:
            print(f"✓ Found {len(groups)} groups.")
            print("\nGroups:")
            for i, group in enumerate(groups[:5], 1):  # Show first 5 groups
                print(f"  {i}. {group['displayName']} (ID: {group['id']})")
            
            if len(groups) > 5:
                print(f"  ... and {len(groups) - 5} more groups")
    except Exception as e:
        print(f"✗ Error getting groups: {str(e)}")
    
    print("\nLooking for Plans...")
    try:
        # Try direct approach
        plans_data = api_func("planner/plans")
        plans = plans_data.get("value", [])
        
        if plans:
            print(f"✓ Found {len(plans)} plans directly.")
            print("\nPlans:")
            for i, plan in enumerate(plans, 1):
                print(f"  {i}. {plan['title']} (ID: {plan['id']})")
            return
    except Exception as e:
        print(f"✗ Direct plan retrieval failed: {str(e)}")
    
    # If direct approach failed, try the group-by-group approach
    print("\nTrying to find plans via groups...")
    all_plans = []
    for group in groups[:5]:  # Limit to first 5 groups
        try:
            group_id = group["id"]
            print(f"  Checking plans for group: {group['displayName']}...")
            group_plans_data = api_func(f"groups/{group_id}/planner/plans")
            group_plans = group_plans_data.get("value", [])
            if group_plans:
                print(f"    ✓ Found {len(group_plans)} plans.")
                all_plans.extend(group_plans)
            else:
                print("    ✗ No plans found.")
        except Exception as e:
            print(f"    ✗ Error: {str(e)}")
    
    if all_plans:
        print(f"\n✓ Found {len(all_plans)} plans via groups.")
        print("\nPlans:")
        for i, plan in enumerate(all_plans, 1):
            print(f"  {i}. {plan['title']} (ID: {plan['id']})")
    else:
        print("\n✗ No plans found through any method.")
        print("\nTroubleshooting suggestions:")
        print("1. Verify API permissions in Azure: Tasks.ReadWrite.All and Group.Read.All")
        print("2. Grant admin consent for these permissions")
        print("3. Check that Microsoft Planner is enabled for your tenant")
        print("4. Create a plan using the Microsoft Planner web interface")

def main():
    """
    Simple script to find and display Microsoft Planner plans
    """
    # Set up argument parser
    parser = argparse.ArgumentParser(description='Find Microsoft Planner plans')
    parser.add_argument('--token', '-t', help='Directly provide an access token instead of using environment variables')
    args = parser.parse_args()
    
    # Load environment variables
    load_dotenv()
    
    # If token is provided via command line, use it directly
    if args.token:
        find_plans_with_token(args.token)
        return
    
    # If token is provided in .env file, use it
    token_from_env = os.getenv("MS_TOKEN_ID")
    if token_from_env:
        print("Using MS_TOKEN_ID from .env file")
        find_plans_with_token(token_from_env)
        return
    
    # Otherwise use the environment variable client credentials approach
    # Check environment variables
    client_id = os.getenv("MS_CLIENT_ID")
    client_secret = os.getenv("MS_CLIENT_SECRET")
    tenant_id = os.getenv("MS_TENANT_ID")
    
    if not all([client_id, client_secret, tenant_id]):
        print("ERROR: Missing required environment variables.")
        print("Please check your .env file has the following variables:")
        print("MS_CLIENT_ID=your_client_id")
        print("MS_CLIENT_SECRET=your_client_secret")
        print("MS_TENANT_ID=your_tenant_id")
        print("\nAlternatively, provide an access token directly using: --token or -t")
        print("Or add MS_TOKEN_ID to your .env file")
        return
        
    print("Checking API connection...")
    try:
        # Test connection
        org_data = call_graph_api("organization")
        print(f"✓ API connection successful to tenant: {org_data['value'][0]['displayName']}")
    except Exception as e:
        print(f"✗ API connection failed: {str(e)}")
        return
    
    # Use the implementation function with the standard call_graph_api
    find_plans_implementation(call_graph_api)

if __name__ == "__main__":
    main()