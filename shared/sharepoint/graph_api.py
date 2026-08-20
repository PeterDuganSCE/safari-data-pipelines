#!/usr/bin/env python
# coding: utf-8


""" 
graph_api.py

Interfaces with the Microsoft Graph API to make various
requests to Microsoft Cloud service resources.
"""

import pandas as pd
import json
import getpass
import requests
import msal
from urllib.parse import urlparse
import io
import sys
import traceback
#One time run per system
####pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org pip-system-certs



########## Global variables to store authentication state ##########
# Sometimes token may expire in the middle of the run
# Store headers here to refresh token if needed
_headers = None



########## Function definitions ##########
"""
@function: authenticate_user

Create and configure Public Client Application with your specific scopes.
Fetch token using username/password.

Note: This requires app to be configured for ROPC (Resource Owner Password
Credentials) flow

@return: header data with token
"""
def authenticate_user():

    ### First step requires us to create a Public Client Application
    # Fetch app registration details
    app_details = json.load(
        open("config/sharepoint_auth.json")
    )['graphAuth']
    
    tenant_id = app_details['tenant_id']
    client_id = app_details['client_id']

    # Create authority URL and define scope(s)
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    
    # Create PublicClientApplication
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority
    )


    ### Next, proceed to fetch access token to use with Graph API
    # Fetch user credentials
    creds = json.load(
        open("C:/Users/" 
            + getpass.getuser() 
            + "/auth.json"))['windowsAuth']

    username = creds['email']
    password = creds['password']
    
    scope = ['https://graph.microsoft.com/.default']
    
    try:
        result = app.acquire_token_by_username_password(
            username = username,
            password = password,
            scopes = scope
        )
        
        if "access_token" in result:
            print("Authentication successful")
            
            token_type = result['token_type']
            token = result['access_token']
            headers = {
                'Authorization': f'{token_type} {token}',
                'Content-Type': 'application/json'
            }
            
            # Update global _headers value
            global _headers
            _headers = headers

            user_data = call_graph_api('me','GET')
            print(f"Authenticated as: {user_data.get('displayName')} ({user_data.get('userPrincipalName')})")

            return headers
        

        else:
            print("Authentication failed")
            print("Error:", result.get('error'))
            print("Error description:", result.get('error_description'))

            return None

            
    except Exception as e:
        print(f"Authentication error: {e}")

        return None


def is_authenticated():
    """
    Returns True when a cached Authorization header is available.
    """
    return bool(_headers and _headers.get("Authorization"))


def ensure_authenticated(force=False):
    """
    Authenticates only when needed (or when force=True).

    Returns:
        bool: True when authenticated, otherwise False.
    """
    if force or not is_authenticated():
        return authenticate_user() is not None

    return True
    


"""
@function: call_graph_api

Sends API calls to Microsoft Graph

@param: endpoint: request to send to Microsoft Graph API
@param: method: HTTP method to use (options: GET, POST, PATCH, or DELETE)
@param: data: data to include in the request body (for POST requests)

@return: response received from Microsoft Graph API call
"""
def call_graph_api(endpoint, method, data=None):
    url = f"https://graph.microsoft.com/v1.0/{endpoint}"
    
    try:
        # Send appropriate request based on HTTP method
        if method == 'GET':
            response = requests.get(url, headers=_headers)
        elif method == 'POST':
            response = requests.post(url, headers=_headers, json=data)
        elif method == 'PATCH':
            response = requests.patch(url, headers=_headers, json=data)
        elif method == 'DELETE':
            response = requests.delete(url, headers=_headers)
        else:
            print("Invalid HTTP method specified")
            return None

        if response.status_code in [200, 201, 202]:
            return response.json() if response.text else {}
        elif response.status_code == 204:
            # Graph DELETE commonly returns 204 No Content.
            return {}
        
        elif response.status_code == 401:
            print(f"Authentication failed (401), attempting token refresh.")

            # Attempt to refresh token if 401 Unauthorized error
            new_headers = authenticate_user()

            # If we were able to refresh token, retry the request   
            if new_headers:
                print('Token refreshed successfully.')
                print('Retrying request with refreshed token.')
                return call_graph_api(endpoint, method, data)  # added the return statement here to return the result of the retried API call; this was recommended by the Copilot code review and is necessary to ensure that the function returns the API response after refreshing the token, rather than returning None after refreshing.

            else:
                print('Failed to refresh token')

                return None
        
        else:
            print(f"API call failed: {response.status_code}")
            print("Response:", response.text)

            return None
        
    except Exception as e:
        print(f"API call error: {e}")

        return None
    


"""
@function: pull_sharepoint_list_data

Pulls SharePoint list data and saves to a dataframe

@param: site_url: SharePoint site where list is housed
@param: list_name: Name of SharePoint list

@return: dataFrame containing SharePoint list data
"""
def pull_sharepoint_list_data(site_url, list_name):
    parsed_site_url = urlparse(site_url)
    # parsed_site_url = urlparse('https://edisonintl.sharepoint.com/teams/td3/AES/AADS/FA/Failure-Event-Analysis')
    host_name = parsed_site_url.netloc
    site_path = parsed_site_url.path

#Was troubleshooting by seeing response
    # response = call_graph_api(f"sites/{host_name}:{site_path}", 'GET')
    # if response is None:
    #     print("Error: Unable to fetch site ID. The API response is None.")
    # else:
    # # Check if 'id' is in the response before accessing it
    #     if 'id' in response:
    #         site_id = response['id']
    #     else:
    #         print("Error: 'id' not found in the API response.")
    #         print("Response:", response)

    # Get site ID associated to SharePoint site
    site_id = call_graph_api(f"sites/{host_name}:{site_path}",'GET')['id']

    # Identify all lists in SharePoint site
    site_lists_json = call_graph_api(f"sites/{site_id}/lists",'GET')
    site_lists = site_lists_json.get('value', [])
    list_lookup = {list['displayName'].upper().strip(): list['id'] for list in site_lists}

    list_id = list_lookup.get(list_name.upper().strip())

    # If list not found in SharePoint site, show user all available lists
    if not list_id:
        print(f"List '{list_name}' not found in site '{site_url}'.")
        print(f"Available lists: {list(list_lookup.keys())}")

        return pd.DataFrame()
    
    else:
        # Create the endpoint to pass to API call
        # Using top here to increase page size from default 200 to speed things up
        endpoint = f"sites/{site_id}/lists/{list_id}/items?$expand=fields&$top=5000"

        # Now begin fetching all records from SharePoint list
        all_list_items = []
        next_link = endpoint
        
        # Handle pagination
        while next_link:
            # Extract just the endpoint part if it's a full URL
            if next_link.startswith("https://graph.microsoft.com/v1.0/"):
                next_link = next_link.replace("https://graph.microsoft.com/v1.0/", "")

            # Pull records for current page
            response = call_graph_api(next_link, 'GET')

            # Append to running list
            all_list_items.extend(response.get('value', []))

            # Update next link for pagination
            next_link = response.get('@odata.nextLink')

            print(len(all_list_items), "items retrieved so far...")


        # Convert to pandas dataframe
        if all_list_items:
            print(f"Retrieved a total of {len(all_list_items)} items")
            
            return pd.json_normalize(all_list_items)
        
        else:
            print("No items found")

            return pd.DataFrame()
    


"""
@function: pull_sharepoint_excel_file

Pulls SharePoint Excel file and saves to a dataframe

@param: site_path: site path where file is housed (e.g., '/teams/ASPDataVisualization')
@param: drive_name: drive or document library name (e.g., 'DV Stakeholders)
@param: file_path: path within the drive (e.g., '/Dashboards/Wire Down/')
@param: file_name: name of the Excel file (e.g., 'Wire Down Flags.xlsx')
@param: sheet_name: name of the sheet in the Excel file (OPTIONAL)

@return: dataFrame containing SharePoint Excel data
"""
def pull_sharepoint_excel_file(site_path, drive_name, file_path, file_name, sheet_name = None):
    host_name = "edisonintl.sharepoint.com"
    
    # Get site ID associated to SharePoint site
    site_id = call_graph_api(f"sites/{host_name}:/{site_path.strip('/')}",'GET')['id']

    # Get all drives for the site to find the specific drive 
    drives_json = call_graph_api(f"sites/{site_id}/drives", 'GET')
    drive_list = drives_json.get('value', [])
    drives_lookup = {drive['name'].upper().strip(): drive['id'] for drive in drive_list}

    drive_id = drives_lookup.get(drive_name.upper().strip())

    # If drive not found in SharePoint site, show user all available drives
    if not drive_id:
        print(f"Drive '{drive_name}' not found in site '{site_path}'.")
        print(f"Available drives: {list(drives_lookup.keys())}")

        return pd.DataFrame()
    
    else:
        # Construct full file path
        # Remove leading slash from file_path if present and ensure file_path ends with /
        full_file_path = file_path.strip('/') + '/' + file_name

        # Get the file by path within the specific drive
        # URL econde the file path for special characters 
        encoded_file_path = requests.utils.quote(full_file_path)
        file_info = call_graph_api(f"sites/{site_id}/drives/{drive_id}/root:/{encoded_file_path}", 'GET')
        
        if not file_info:
            print(f"File not found: {full_file_path} in drive '{drive_name}'")

            return pd.DataFrame()
        
        # Get the download URL for the file content
        download_url = file_info.get('@microsoft.graph.downloadUrl')

        # Download the file content
        response = requests.get(download_url)

        if response.status_code == 200:
            # Read the Excel file from bytes
            excel_data = io.BytesIO(response.content)
            
            # Read into pandas DataFrame
            if sheet_name:
                df = pd.read_excel(excel_data, sheet_name=sheet_name)
                
            else:
                df = pd.read_excel(excel_data)  # Reads first sheet by default
            
            return df
        
        else:
            print(f"Failed to download file: {response.status_code}")

            return pd.DataFrame()



"""
@function: send_email

Send an email using Microsoft Graph API

@param: subject: email subject line
@param: body: email body content
@param: to_recipients: list of email addresses or single email address
@param: cc_recipients: list of CC email addresses (OPTIONAL)
@param: bcc_recipients: list of BCC email addresses (OPTIONAL)
"""
def send_email(subject, body, 
               to_recipients, cc_recipients=None,bcc_recipients=None):

    # Ensure to_recipients is a list
    if isinstance(to_recipients, str):
        to_recipients = [to_recipients]
    
    # Build recipient lists
    to_list = [{"emailAddress": {"address": email}} for email in to_recipients]
    cc_list = [{"emailAddress": {"address": email}} for email in (cc_recipients or [])]
    bcc_list = [{"emailAddress": {"address": email}} for email in (bcc_recipients or [])]
    
    # Prepare email message
    message = {
        "subject": subject,
        "body": {
            "contentType": "Text",
            "content": body
        },
        "toRecipients": to_list,
    }
    
    # Add CC recipients if specified
    if cc_list:
        message["ccRecipients"] = cc_list
    
    # Add BCC recipients if specified
    if bcc_list:
        message["bccRecipients"] = bcc_list
    
    # Create the request payload
    email_data = {
        "message": message
    }
    
    
    # Send email
    call_graph_api("me/sendMail", 'POST', data=email_data)
    

def resolve_site_id_from_url(site_url):
    """
    Resolve a SharePoint site URL to the Microsoft Graph site ID.
    """
    parsed_site_url = urlparse(site_url)
    host_name = parsed_site_url.netloc
    site_path = parsed_site_url.path

    if not host_name or not site_path:
        print(f"Invalid site URL: {site_url}")
        return None

    result = call_graph_api(f"sites/{host_name}:{site_path}", "GET")
    if not result or "id" not in result:
        print("Failed to resolve SharePoint site ID from URL.")
        return None

    return result["id"]


def get_drive_id(site_id, drive_name="Documents"):
    endpoint = f"sites/{site_id}/drives"
    result = call_graph_api(endpoint, "GET")

    if not result or "value" not in result:
        return None

    for drive in result["value"]:
        if drive["name"] == drive_name:
            return drive["id"]

    print(f"Drive '{drive_name}' not found")
    return None


def create_sharepoint_folder(site_id, drive_id, parent_path, folder_name):
    """
    Creates a folder in a SharePoint document library using Microsoft Graph.

    Parameters:
        site_id (str): The SharePoint site ID.
        drive_id (str): The document library drive ID.
        parent_path (str): The path in the document library where the folder should be created.
                           Example: "Shared Documents/Reports"
                           Use "" or None to create at the root of the drive.
        folder_name (str): The name of the folder to create.

    Returns:
        dict or None: The Graph API response if successful, otherwise None.
    """

    # Get drive ID for the document library (if not already provided)
    if not drive_id:
        drive_id = get_drive_id(site_id)
        if not drive_id:
            print("Failed to get drive ID. Cannot create folder.")
            return None

    # Payload required by Microsoft Graph to create a folder
    data = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename"  # Options: "fail", "replace", "rename"
    }

    # If creating folder at root of the document library
    if not parent_path:
        endpoint = f"sites/{site_id}/drives/{drive_id}/root/children"
    else:
        # Clean up leading/trailing slashes
        parent_path = parent_path.strip("/")

        endpoint = (
            f"sites/{site_id}/drives/{drive_id}"
            f"/root:/{parent_path}:/children"
        )

    return call_graph_api(
        endpoint=endpoint,
        method="POST",
        data=data
    )


def delete_sharePoint_folder(site_id, drive_id=None, parent_path=None, folder_name=None, folder_item_id=None):
    """
    Deletes a folder in a SharePoint document library using Microsoft Graph.

    Parameters:
        site_id (str): The SharePoint site ID.
        drive_id (str | None): The document library drive ID. If omitted, defaults to "Documents".
        parent_path (str | None): Parent folder path for path-based deletion.
        folder_name (str | None): Folder name for path-based deletion.
        folder_item_id (str | None): Folder item ID for ID-based deletion.

    Returns:
        bool: True if folder delete request succeeded, otherwise False.
    """

    if not site_id:
        print("site_id is required.")
        return False

    # Get drive ID for the document library (if not already provided)
    if not drive_id:
        drive_id = get_drive_id(site_id)
        if not drive_id:
            print("Failed to get drive ID. Cannot delete folder.")
            return False

    if folder_item_id:
        endpoint = f"sites/{site_id}/drives/{drive_id}/items/{folder_item_id}"
    else:
        if not folder_name:
            print("Provide either folder_item_id, or folder_name (+ optional parent_path).")
            return False

        if parent_path:
            parent_path = parent_path.strip("/")
            folder_path = f"{parent_path}/{folder_name}" if parent_path else folder_name
        else:
            folder_path = folder_name

        endpoint = f"sites/{site_id}/drives/{drive_id}/root:/{folder_path.strip('/')}"

    result = call_graph_api(endpoint=endpoint, method="DELETE")
    if result is None:
        print("Failed to delete folder.")
        return False

    return True


def rename_sharePoint_folder(site_id, new_folder_name, drive_id=None, parent_path=None, old_folder_name=None, folder_item_id=None):
    """
    Renames a folder in a SharePoint document library using Microsoft Graph.

    Parameters:
        site_id (str): The SharePoint site ID.
        new_folder_name (str): New name for the folder.
        drive_id (str | None): The document library drive ID. If omitted, defaults to "Documents".
        parent_path (str | None): Parent folder path for path-based rename.
        old_folder_name (str | None): Existing folder name for path-based rename.
        folder_item_id (str | None): Folder item ID for ID-based rename.

    Returns:
        dict | None: Updated folder metadata when successful, otherwise None.
    """

    if not site_id:
        print("site_id is required.")
        return None

    if not new_folder_name:
        print("new_folder_name is required.")
        return None

    # Get drive ID for the document library (if not already provided)
    if not drive_id:
        drive_id = get_drive_id(site_id)
        if not drive_id:
            print("Failed to get drive ID. Cannot rename folder.")
            return None

    if folder_item_id:
        endpoint = f"sites/{site_id}/drives/{drive_id}/items/{folder_item_id}"
    else:
        if not old_folder_name:
            print("Provide either folder_item_id, or old_folder_name (+ optional parent_path).")
            return None

        if parent_path:
            parent_path = parent_path.strip("/")
            folder_path = f"{parent_path}/{old_folder_name}" if parent_path else old_folder_name
        else:
            folder_path = old_folder_name

        endpoint = f"sites/{site_id}/drives/{drive_id}/root:/{folder_path.strip('/')}"

    data = {
        "name": new_folder_name
    }

    result = call_graph_api(endpoint=endpoint, method="PATCH", data=data)
    if result is None:
        print("Failed to rename folder.")
        return None

    return result


########## Main execution block ##########
# if __name__ == "__main__":
#     try:
#         ##### First authenticate
#         authenticate_user()


#         ###### Some example API calls
#         ### [EXAMPLE 1] Pulling data from SharePoint lists
#         # site_url = 'https://edisonintl.sharepoint.com/teams/ASPDataVisualization'
#         # list_name = 'EFD Risk Scores'
        
#         site_url = 'https://edisonintl.sharepoint.com/teams/td3/AES/AADS/FA/Failure-Event-Analysis'
#         list_name = 'Repair Order Events'

#         # Pull all items in the list
#         list_df = pull_sharepoint_list_data(site_url, list_name)


#         ### [EXAMPLE 2] Pulling data from SharePoint Excel files
#         site_path = '/teams/ASPDataVisualization'
#         drive_name = 'DV Stakeholders'
#         file_path = '/Dashboards/Wire Down/'
#         file_name = 'Wire Down Flags.xlsx'
#         sheet_name = 'Wire Down Flags'

#         # Pull data from the Excel file
#         excel_df = pull_sharepoint_excel_file(site_path, drive_name, file_path, file_name, sheet_name = sheet_name)


#         ### [EXAMPLE 3] Interfacing with Planner data
#         plan_id = 'iSR7pGASK0WBmTVb9qsHqmQAFurd'    # Replace with your specific plan ID
        
#         # Get all buckets for a plan
#         buckets_json = call_graph_api(f"planner/plans/{plan_id}/buckets",'GET')
#         buckets = buckets_json.get('value', []) if buckets_json else []
#         bucket_lookup = {bucket['id']: bucket['name'] for bucket in buckets}

#         # Get all tasks for a plan
#         tasks_json = call_graph_api(f"planner/plans/{plan_id}/tasks",'GET')
#         tasks = tasks_json.get('value', []) 

#         # Convert to pandas dataframe
#         tasks_df = pd.DataFrame(tasks)

#         # Lookup bucket name associated with bucket ID and add to dataframe
#         tasks_df['bucketName'] = tasks_df['bucketId'].map(bucket_lookup)


#         ### [EXAMPLE 4] Sending emails
#         subject = 'insertProjectNameHere Results'
#         body = \
#         """
#         The Python job ran successfully to completion.

#         You can insert any other text you would like to include here as well.
#         """
#         to_recipients = ['insertUserEmailHere@sce.com']     # If multiple recipients, separate with a comma

#         send_email(
#             subject,
#             body,
#             to_recipients
#         )


#     except Exception as e:
#         ##### If error encountered, send email alert
#         # Fetch error and save to variable so we can include in email
#         error_type, error_value, error_traceback = sys.exc_info()

#         tb_str = ''.join(traceback.format_tb(error_traceback)) 
#         error_desc = f'An error occurred: {error_type}: {error_value}\n\nTraceback:\n{tb_str}' 
    
#         # Send email
#         subject = 'insertProjectNameHere Results'
#         body = error_desc
#         to_recipients = ['insertUserEmailHere@sce.com']     # If multiple recipients, separate with a comma

#         send_email(
#             subject,
#             body,
#             to_recipients
#         )