
import sys
import traceback
import json
import logging
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

import pandas as pd
import requests
from . import graph_api

try:
    from control.controller import Controller
except ModuleNotFoundError:
    Controller = None

from .graph_api import (
    create_sharepoint_folder,
    delete_sharePoint_folder,
    ensure_authenticated,
    get_drive_id,
    rename_sharePoint_folder,
    resolve_site_id_from_url,
    send_email
)


site_url = 'https://edisonintl.sharepoint.com/teams/td3/AES/AADS/FA/Failure-Event-Analysis'
drive_id = None  # Set to None to fetch drive ID dynamically based on site URL and drive name

logger = Controller.get_controller().logger if Controller else logging.getLogger(__name__)


def upload_dataframe_to_sharepoint_list(
    sharepoint_site: str,
    sharepoint_list_name: str,
    dataframe: pd.DataFrame,
    sharepoint_field_mapping: dict[str, str] | None = None,
) -> int:
    """Upload each DataFrame row as an item in a SharePoint list.

    Args:
        sharepoint_site: Full SharePoint site URL.
        sharepoint_list_name: Display name of the destination SharePoint list.
        dataframe: Rows to upload.
        sharepoint_field_mapping: Optional mapping from DataFrame column names
            to SharePoint internal field names.

    Returns:
        Number of rows uploaded.
    """
    if not sharepoint_list_name:
        raise ValueError("sharepoint_list_name must be provided")

    if not isinstance(dataframe, pd.DataFrame):
        raise TypeError("dataframe must be a pandas DataFrame")

    ensure_authenticated()
    field_mapping = sharepoint_field_mapping or {}
    parsed_site_url = urlparse(sharepoint_site)
    if not parsed_site_url.netloc or not parsed_site_url.path:
        raise ValueError(f"Invalid SharePoint site URL: {sharepoint_site}")

    site_response = graph_api.call_graph_api(
        f"sites/{parsed_site_url.netloc}:{parsed_site_url.path}",
        "GET",
    )
    if not site_response or "id" not in site_response:
        raise RuntimeError(f"Unable to resolve SharePoint site: {sharepoint_site}")

    site_id = site_response["id"]
    lists_response = graph_api.call_graph_api(f"sites/{site_id}/lists", "GET") or {}
    list_lookup = {
        item["displayName"].strip().casefold(): item["id"]
        for item in lists_response.get("value", [])
        if item.get("displayName") and item.get("id")
    }
    list_id = list_lookup.get(sharepoint_list_name.strip().casefold())
    if not list_id:
        raise ValueError(
            f"SharePoint list '{sharepoint_list_name}' was not found. "
            f"Available lists: {sorted(list_lookup)}"
        )

    def serialize_value(value):
        if pd.isna(value):
            return None
        if isinstance(value, (pd.Timestamp, datetime)):
            return value.isoformat()
        if hasattr(value, "item"):
            return value.item()
        return value

    for _, row in dataframe.iterrows():
        fields = {
            field_mapping.get(column, column): serialize_value(value)
            for column, value in row.items()
        }
        response = graph_api.call_graph_api(
            f"sites/{site_id}/lists/{list_id}/items",
            "POST",
            data={"fields": fields},
        )
        if not response:
            raise RuntimeError(
                f"SharePoint list upload failed for row with fields: {fields}"
            )

    logger.info(
        "Uploaded %s rows to SharePoint list %s.",
        len(dataframe),
        sharepoint_list_name,
    )
    return len(dataframe)


def _log_file_only(level: int, message: str) -> None:
    """Write a log message to non-console logger handlers only."""
    for handler in logger.handlers:
        if isinstance(handler, logging.FileHandler):
            record = logger.makeRecord(
                logger.name,
                level,
                __file__,
                0,
                message,
                args=(),
                exc_info=None
            )
            handler.handle(record)


def _normalize_drive_name(name: str) -> str:
    """Normalize a SharePoint library name for loose matching."""
    return " ".join((name or "").replace("%20", " ").split()).strip().casefold()


def _resolve_site_id_silently(candidate_site_url: str) -> str | None:
    """Resolve a site ID without emitting expected probe failures to console."""
    parsed_site_url = urlparse(candidate_site_url)
    host_name = parsed_site_url.netloc
    site_path = parsed_site_url.path

    if not host_name or not site_path:
        return None

    auth_header = (graph_api._headers or {}).get("Authorization")
    if not auth_header:
        if not ensure_authenticated(force=True):
            return None
        auth_header = (graph_api._headers or {}).get("Authorization")

    headers = {
        "Authorization": auth_header,
        "Content-Type": "application/json"
    }
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{host_name}:{site_path}"
    response = requests.get(endpoint, headers=headers)

    if response.status_code == 401 and ensure_authenticated(force=True):
        auth_header = (graph_api._headers or {}).get("Authorization")
        headers["Authorization"] = auth_header
        response = requests.get(endpoint, headers=headers)

    if response.status_code == 200:
        result = response.json() if response.text else {}
        return result.get("id")

    if response.status_code == 404:
        _log_file_only(logging.INFO, f"SharePoint site probe miss: {candidate_site_url}")

    return None


def _resolve_sharepoint_upload_target(sharepoint_folder: str) -> tuple[str, str, str] | None:
    """Resolve a SharePoint folder URL or path into site ID, drive ID, and folder path."""
    if not sharepoint_folder:
        default_site_id = resolve_site_id_from_url(site_url)
        if not default_site_id:
            return None

        default_drive_id = get_drive_id(default_site_id)
        if not default_drive_id:
            return None

        return default_site_id, default_drive_id, ""

    folder_value = sharepoint_folder.strip()
    parsed_url = urlparse(folder_value)

    if not (parsed_url.scheme and parsed_url.netloc):
        default_site_id = resolve_site_id_from_url(site_url)
        if not default_site_id:
            return None

        path_segments = [segment for segment in folder_value.replace("\\", "/").split("/") if segment]
        drives_result = graph_api.call_graph_api(f"sites/{default_site_id}/drives", "GET")
        drives = (drives_result or {}).get("value", [])
        if path_segments:
            requested_drive = _normalize_drive_name(path_segments[0])
            for drive in drives:
                if requested_drive == _normalize_drive_name(drive.get("name", "")):
                    return default_site_id, drive["id"], "/".join(path_segments[1:])

        default_drive_id = get_drive_id(default_site_id)
        if not default_drive_id:
            return None

        return default_site_id, default_drive_id, _normalize_sharepoint_folder_path(folder_value)

    decoded_segments = [segment for segment in unquote(parsed_url.path).split("/") if segment]
    if not decoded_segments:
        return None

    site_id = None
    remaining_segments: list[str] = []
    for end_index in range(len(decoded_segments), 0, -1):
        candidate_site_url = f"{parsed_url.scheme}://{parsed_url.netloc}/{'/'.join(decoded_segments[:end_index])}"
        candidate_site_id = _resolve_site_id_silently(candidate_site_url)
        if candidate_site_id:
            site_id = candidate_site_id
            remaining_segments = decoded_segments[end_index:]
            break

    if not site_id:
        logger.error(f"Unable to resolve site from SharePoint URL: {sharepoint_folder}")
        return None

    if not remaining_segments:
        drive_id = get_drive_id(site_id)
        if not drive_id:
            return None
        return site_id, drive_id, ""

    drives_result = graph_api.call_graph_api(f"sites/{site_id}/drives", "GET")
    drives = (drives_result or {}).get("value", [])
    if not drives:
        logger.error(f"Unable to load drives for resolved site: {sharepoint_folder}")
        return None

    library_segment = remaining_segments[0]
    normalized_library = _normalize_drive_name(library_segment)
    drive_match = None

    for drive in drives:
        drive_name = drive.get("name", "")
        normalized_drive_name = _normalize_drive_name(drive_name)
        if normalized_library == normalized_drive_name:
            drive_match = drive
            break

        if normalized_library == "shared documents" and normalized_drive_name == "documents":
            drive_match = drive
            break

    if not drive_match:
        logger.error(f"Unable to match SharePoint library '{library_segment}' for URL: {sharepoint_folder}")
        return None

    folder_path = "/".join(remaining_segments[1:]).strip("/")
    return site_id, drive_match["id"], folder_path


def _normalize_sharepoint_folder_path(sharepoint_folder: str) -> str:
    """Convert a SharePoint folder URL or path into a drive-relative folder path."""
    if not sharepoint_folder:
        return ""

    folder_value = sharepoint_folder.strip()
    parsed_url = urlparse(folder_value)

    if parsed_url.scheme and parsed_url.netloc:
        folder_value = unquote(parsed_url.path).strip("/")

        site_path = urlparse(site_url).path.strip("/")
        if folder_value.startswith(site_path):
            folder_value = folder_value[len(site_path):].strip("/")

    folder_value = folder_value.strip("/")

    # The resolved drive ID points at the default Documents library, so paths
    # passed to Graph must be relative to that drive root rather than including
    # the library name from a browser URL.
    if folder_value.startswith("Shared Documents/"):
        folder_value = folder_value[len("Shared Documents/"):]
    elif folder_value == "Shared Documents":
        folder_value = ""

    return folder_value

def create_folder(folder_name: str, path_name: str) -> str | None:
    """Create a folder in SharePoint.

    Args:
        folder_name (str): Name of the folder to create.
        path_name (str): Parent SharePoint path where the folder will be created.

    Returns:
        str | None: The created folder's SharePoint web URL when successful;
            otherwise None.
    """
    try:
        if not ensure_authenticated():
            print("Authentication failed. Exiting.")
            return None

        # Resolve tenant/site context from the configured SharePoint URL.
        site_id = resolve_site_id_from_url(site_url)

        if not site_id:
            print("Unable to resolve site ID. Exiting.")
            # sys.exit(1)
        
        # Resolve the document library (drive) for folder operations.
        drive_id = get_drive_id(site_id)
                
        if not drive_id:
            print("Failed to get drive ID. Exiting.")
            # sys.exit(1)

        print(f"Site ID: {site_id}")
        print(f"Drive ID: {drive_id}")
        
        # For a site's default 'Documents' library, create in drive root by default.
        result = create_sharepoint_folder(site_id, drive_id, parent_path=path_name, folder_name=folder_name)
        print("API call result:")
        print(json.dumps(result, indent=4) if result else result)

        if result:
            return result["webUrl"]
        
    except Exception as e:
        ##### If error encountered, send email alert
        # Fetch error and save to variable so we can include in email
        error_type, error_value, error_traceback = sys.exc_info()
    
        tb_str = ''.join(traceback.format_tb(error_traceback)) 
        error_desc = f'An error occurred: {error_type}: {error_value}\n\nTraceback:\n{tb_str}' 
        
        # Send email
        subject = 'insertProjectNameHere Results'
        body = error_desc
        to_recipients = ['peter.dugan@sce.com']     # If multiple recipients, separate with a comma
    
        send_email(
            subject,
            body,
            to_recipients
        )

         
def delete_folder(folder_name: str, path_name: str) -> bool:
    """Delete a folder in SharePoint.

    Args:
        folder_name (str): Name of the folder to delete.
        path_name (str): Parent SharePoint path that contains the folder.

    Returns:
        bool: True if the folder is deleted successfully; otherwise False.
    """
    try:
        if not ensure_authenticated():
            print("Authentication failed. Exiting.")
            return False

        # Resolve tenant/site context from the configured SharePoint URL.
        site_id = resolve_site_id_from_url(site_url)

        if not site_id:
            print("Unable to resolve site ID. Exiting.")
            return False

        # Resolve the document library (drive) for folder operations.
        drive_id = get_drive_id(site_id)

        if not drive_id:
            print("Failed to get drive ID. Exiting.")
            return False

        print(f"Site ID: {site_id}")
        print(f"Drive ID: {drive_id}")

        result = delete_sharePoint_folder(
            site_id=site_id,
            drive_id=drive_id,
            parent_path=path_name,
            folder_name=folder_name
        )

        if result:
            print(f"Deleted folder '{folder_name}'.")
            return True

        return False

    except Exception as e:
        ##### If error encountered, send email alert
        # Fetch error and save to variable so we can include in email
        error_type, error_value, error_traceback = sys.exc_info()

        tb_str = ''.join(traceback.format_tb(error_traceback))
        error_desc = f'An error occurred: {error_type}: {error_value}\n\nTraceback:\n{tb_str}'

        # Send email
        subject = 'insertProjectNameHere Results'
        body = error_desc
        to_recipients = ['peter.dugan@sce.com']     # If multiple recipients, separate with a comma

        send_email(
            subject,
            body,
            to_recipients
        )

        return False


def rename_folder(old_name: str, new_name: str, path_name: str) -> bool:
    """Rename a SharePoint folder.

    Args:
        old_name (str): Current folder name.
        new_name (str): New folder name.
        path_name (str): Parent SharePoint path that contains the folder.

    Returns:
        bool: True if the folder is renamed successfully; otherwise False.
    """
    try:
        if not ensure_authenticated():
            print("Authentication failed. Exiting.")
            return False

        # Resolve tenant/site context from the configured SharePoint URL.
        site_id = resolve_site_id_from_url(site_url)

        if not site_id:
            print("Unable to resolve site ID. Exiting.")
            return False

        # Resolve the document library (drive) for folder operations.
        drive_id = get_drive_id(site_id)

        if not drive_id:
            print("Failed to get drive ID. Exiting.")
            return False

        print(f"Site ID: {site_id}")
        print(f"Drive ID: {drive_id}")

        result = rename_sharePoint_folder(
            site_id=site_id,
            drive_id=drive_id,
            parent_path=path_name,
            old_folder_name=old_name,
            new_folder_name=new_name
        )

        if result:
            print(f"Renamed folder '{old_name}' to '{new_name}'.")
            return True

        return False

    except Exception as e:
        ##### If error encountered, send email alert
        # Fetch error and save to variable so we can include in email
        error_type, error_value, error_traceback = sys.exc_info()

        tb_str = ''.join(traceback.format_tb(error_traceback))
        error_desc = f'An error occurred: {error_type}: {error_value}\n\nTraceback:\n{tb_str}'

        # Send email
        subject = 'insertProjectNameHere Results'
        body = error_desc
        to_recipients = ['peter.dugan@sce.com']     # If multiple recipients, separate with a comma

        send_email(
            subject,
            body,
            to_recipients
        )

        return False


def upload_file(filepath: Path, sharepoint_folder: str) -> str | None:
    """Upload a local file to a target SharePoint folder.

    Args:
        filepath (Path): Local file path to upload.
        sharepoint_folder (str): Target SharePoint folder to upload the file to.

    Returns:
        str | None: SharePoint web URL for the uploaded file when successful;
            otherwise None.
    """
    try:
        local_path = Path(filepath)
        if not local_path.exists() or not local_path.is_file():
            logger.error(f"Local file not found: {local_path}")  # print(f"Local file not found: {local_path}")
            return None

        if not ensure_authenticated():
            logger.error("Authentication failed. Exiting.")  # print("Authentication failed. Exiting.")
            return None

        # Resolve tenant/site, document library, and drive-relative folder path.
        upload_target = _resolve_sharepoint_upload_target(sharepoint_folder)
        if not upload_target:
            logger.error("Unable to resolve SharePoint upload target. Exiting.")
            return None

        site_id, drive_id, folder_path = upload_target

        sp_relative_path = f"{folder_path}/{local_path.name}" if folder_path else local_path.name
        encoded_sp_path = quote(sp_relative_path, safe="/")

        # Direct upload endpoint for file content.
        upload_url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}"
            f"/root:/{encoded_sp_path}:/content"
        )

        # Ensure we have a usable Authorization header from the shared auth module.
        auth_header = (graph_api._headers or {}).get("Authorization")
        if not auth_header:
            if not ensure_authenticated(force=True):
                print("Unable to refresh authentication token. Exiting.")
                return None
            auth_header = (graph_api._headers or {}).get("Authorization")

        headers = {
            "Authorization": auth_header,
            "Content-Type": "application/octet-stream"
        }

        with open(local_path, "rb") as file_data:
            response = requests.put(upload_url, headers=headers, data=file_data)

        if response.status_code in (200, 201):
            result = response.json() if response.text else {}
            web_url = result.get("webUrl")
            print(f"Uploaded file '{local_path.name}' to '{sharepoint_folder}'.")
            if web_url:
                print(f"Uploaded file URL: {web_url}")
            return web_url

        # Handle expired token by re-authenticating once and retrying upload.
        if response.status_code == 401 and ensure_authenticated(force=True):
            auth_header = (graph_api._headers or {}).get("Authorization")
            headers["Authorization"] = auth_header
            with open(local_path, "rb") as file_data:
                response = requests.put(upload_url, headers=headers, data=file_data)

            if response.status_code in (200, 201):
                result = response.json() if response.text else {}
                web_url = result.get("webUrl")
                print(f"Uploaded file '{local_path.name}' to '{sharepoint_folder}'.")
                if web_url:
                    print(f"Uploaded file URL: {web_url}")
                return web_url

        error_result = response.text
        try:
            parsed_error = response.json()
            error_result = json.dumps(parsed_error, indent=4)
        except ValueError:
            pass

        print(f"File upload failed: {response.status_code}")
        print(f"Response: {error_result}")

        return None

    except Exception as e:
        ##### If error encountered, send email alert
        # Fetch error and save to variable so we can include in email
        error_type, error_value, error_traceback = sys.exc_info()

        tb_str = ''.join(traceback.format_tb(error_traceback))
        error_desc = f'An error occurred: {error_type}: {error_value}\n\nTraceback:\n{tb_str}'

        # Send email
        subject = 'insertProjectNameHere Results'
        body = error_desc
        to_recipients = ['peter.dugan@sce.com']     # If multiple recipients, separate with a comma

        send_email(
            subject,
            body,
            to_recipients
        )

        return None


def download_file(file_name: str, sharepoint_folder: str, destination: Path) -> Path | None:
    """Download a SharePoint file to a local working path."""
    try:
        if not ensure_authenticated():
            logger.error("Authentication failed. Exiting.")
            return None

        download_target = _resolve_sharepoint_upload_target(sharepoint_folder)
        if not download_target:
            logger.error("Unable to resolve SharePoint download target. Exiting.")
            return None

        site_id, drive_id, folder_path = download_target
        relative_path = f"{folder_path}/{file_name}" if folder_path else file_name
        encoded_path = quote(relative_path, safe="/")
        file_info = graph_api.call_graph_api(
            f"sites/{site_id}/drives/{drive_id}/root:/{encoded_path}",
            "GET"
        )
        download_url = (file_info or {}).get("@microsoft.graph.downloadUrl")
        if not download_url:
            logger.error(f"SharePoint file not found: {relative_path}")
            return None

        response = requests.get(download_url)
        if response.status_code != 200:
            logger.error(f"File download failed: {response.status_code}")
            return None

        destination_path = Path(destination)
        destination_path.write_bytes(response.content)
        return destination_path
    except Exception as e:
        logger.error(f"File download error: {e}")
        return None


def delete_file(filename: str) -> bool:
    """Delete a file from SharePoint.

    Args:
        filename (str): Full SharePoint file URL or drive-relative file path.

    Returns:
        bool: True if the file is deleted successfully; otherwise False.
    """
    try:
        if not filename:
            logger.error("No SharePoint file path was provided.")
            return False

        if not ensure_authenticated():
            logger.error("Authentication failed. Exiting.")
            return False

        target_value = filename.strip()
        parsed_url = urlparse(target_value)

        if parsed_url.scheme and parsed_url.netloc:
            parent_url, _, file_name = target_value.rpartition("/")
            if not file_name:
                logger.error(f"Unable to determine file name from URL: {filename}")
                return False

            upload_target = _resolve_sharepoint_upload_target(parent_url)
            if not upload_target:
                logger.error("Unable to resolve SharePoint file target. Exiting.")
                return False

            site_id, drive_id, folder_path = upload_target
        else:
            normalized_path = target_value.strip("/")
            if not normalized_path:
                logger.error("No SharePoint file path was provided.")
                return False

            path_parts = normalized_path.split("/")
            file_name = path_parts[-1]
            folder_path = "/".join(path_parts[:-1]).strip("/")

            site_id = resolve_site_id_from_url(site_url)
            if not site_id:
                logger.error("Unable to resolve site ID. Exiting.")
                return False

            drive_id = get_drive_id(site_id)
            if not drive_id:
                logger.error("Failed to get drive ID. Exiting.")
                return False

        relative_file_path = f"{folder_path}/{file_name}" if folder_path else file_name
        encoded_file_path = quote(relative_file_path, safe="/")
        endpoint = f"sites/{site_id}/drives/{drive_id}/root:/{encoded_file_path}"

        result = graph_api.call_graph_api(endpoint, "DELETE")
        if result is None:
            logger.error(f"Failed to delete file: {filename}")
            return False

        print(f"Deleted file '{file_name}' from '{filename}'.")
        return True

    except Exception as e:
        ##### If error encountered, send email alert
        # Fetch error and save to variable so we can include in email
        error_type, error_value, error_traceback = sys.exc_info()

        tb_str = ''.join(traceback.format_tb(error_traceback))
        error_desc = f'An error occurred: {error_type}: {error_value}\n\nTraceback:\n{tb_str}'

        # Send email
        subject = 'insertProjectNameHere Results'
        body = error_desc
        to_recipients = ['peter.dugan@sce.com']     # If multiple recipients, separate with a comma

        send_email(
            subject,
            body,
            to_recipients
        )

        return False
