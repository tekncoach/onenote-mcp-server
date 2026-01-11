#!/usr/bin/env python3
"""
OneNote MCP Server

A Model Context Protocol server for Microsoft OneNote integration.
This allows Claude Desktop to read and interact with OneNote notebooks.
"""

import os
import asyncio
import json
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path
import time
from msal import ConfidentialClientApplication, PublicClientApplication
import httpx
from fastmcp import FastMCP

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastMCP instance
mcp = FastMCP("OneNote MCP Server")

# Microsoft Graph API constants
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
SCOPES = [
    "https://graph.microsoft.com/Notes.Read",
    "https://graph.microsoft.com/Notes.ReadWrite",
    "https://graph.microsoft.com/Files.ReadWrite",
    "https://graph.microsoft.com/User.Read"
]

# Token cache configuration
TOKEN_CACHE_ENABLED = os.getenv("ONENOTE_CACHE_TOKENS", "true").lower() in ("true", "1", "yes")
TOKEN_CACHE_FILE = Path.home() / ".onenote_mcp_tokens.json"

# Global variables for authentication
access_token: Optional[str] = None
refresh_token: Optional[str] = None
token_expires_at: Optional[float] = None
msal_app: Optional[PublicClientApplication] = None

def get_client_id() -> str:
    """Get the Azure client ID from environment variable."""
    client_id = os.getenv("AZURE_CLIENT_ID")
    if not client_id:
        raise Exception("AZURE_CLIENT_ID environment variable not set")
    return client_id

def save_tokens(access_tok: str, refresh_tok: str = None, expires_in: int = 3600) -> None:
    """Save tokens to disk for persistence across sessions."""
    global access_token, refresh_token, token_expires_at
    
    access_token = access_tok
    if refresh_tok:
        refresh_token = refresh_tok
    token_expires_at = time.time() + expires_in - 300  # 5 min buffer
    
    # Only save to disk if caching is enabled
    if not TOKEN_CACHE_ENABLED:
        logger.info("Token caching disabled - tokens will not persist across sessions")
        return
    
    try:
        token_data = {
            "access_token": access_token,
            "refresh_token": refresh_token,
            "expires_at": token_expires_at
        }
        
        with open(TOKEN_CACHE_FILE, 'w') as f:
            json.dump(token_data, f)
        
        # Set secure permissions (user read/write only)
        TOKEN_CACHE_FILE.chmod(0o600)
        logger.info(f"Tokens saved to {TOKEN_CACHE_FILE}")
        
    except Exception as e:
        logger.warning(f"Failed to save tokens: {e}")

def load_tokens() -> bool:
    """Load tokens from disk. Returns True if valid tokens loaded."""
    global access_token, refresh_token, token_expires_at
    
    # Don't load tokens if caching is disabled
    if not TOKEN_CACHE_ENABLED:
        logger.info("Token caching disabled - will not load cached tokens")
        return False
    
    try:
        if not TOKEN_CACHE_FILE.exists():
            logger.info(f"No token cache file found at {TOKEN_CACHE_FILE}")
            return False
            
        with open(TOKEN_CACHE_FILE, 'r') as f:
            token_data = json.load(f)
        
        access_token = token_data.get("access_token")
        refresh_token = token_data.get("refresh_token")
        token_expires_at = token_data.get("expires_at")
        
        # Check if token is still valid
        if token_expires_at and time.time() < token_expires_at:
            logger.info(f"Valid tokens loaded from {TOKEN_CACHE_FILE}")
            return True
        else:
            logger.info("Cached tokens expired")
            return False
            
    except Exception as e:
        logger.warning(f"Failed to load tokens: {e}")
        return False

async def refresh_access_token() -> bool:
    """Try to refresh the access token using the refresh token."""
    global access_token, msal_app
    
    if not refresh_token or not msal_app:
        return False
    
    try:
        # Try to get accounts from MSAL cache
        accounts = msal_app.get_accounts()
        
        if accounts:
            # Try silent acquisition first
            result = msal_app.acquire_token_silent(SCOPES, account=accounts[0])
            
            if result and "access_token" in result:
                save_tokens(
                    result["access_token"],
                    result.get("refresh_token", refresh_token),
                    result.get("expires_in", 3600)
                )
                logger.info("Token refreshed successfully via MSAL silent acquisition")
                return True
        
        # MSAL silent acquisition failed - try manual refresh with cached refresh token
        logger.info("MSAL silent acquisition failed, trying manual refresh with cached token")
        return await manual_token_refresh()
        
    except Exception as e:
        logger.warning(f"Token refresh error: {e}")
        return False

async def manual_token_refresh() -> bool:
    """Manually refresh access token using cached refresh token."""
    global access_token, refresh_token
    
    if not refresh_token:
        logger.info("No refresh token available for manual refresh")
        return False
    
    try:
        client_id = get_client_id()
        
        # Microsoft token endpoint
        token_url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
        
        # Prepare refresh token request
        data = {
            "grant_type": "refresh_token",
            "refresh_token": refresh_token,
            "client_id": client_id,
            "scope": " ".join(SCOPES + ["offline_access"])  # Include offline_access for refresh requests
        }
        
        # Make the refresh request
        async with httpx.AsyncClient() as client:
            response = await client.post(
                token_url,
                data=data,
                headers={"Content-Type": "application/x-www-form-urlencoded"}
            )
            
            if response.status_code == 200:
                token_data = response.json()
                
                # Save the new tokens
                save_tokens(
                    token_data["access_token"],
                    token_data.get("refresh_token", refresh_token),  # Use new refresh token if provided
                    token_data.get("expires_in", 3600)
                )
                
                logger.info("Token refreshed successfully via manual refresh")
                return True
            else:
                logger.warning(f"Manual token refresh failed: {response.status_code} - {response.text}")
                return False
                
    except Exception as e:
        logger.warning(f"Manual token refresh error: {e}")
        return False

def init_msal_app(client_id: str) -> PublicClientApplication:
    """Initialize MSAL application for authentication."""
    # Create a simple in-memory cache for MSAL
    return PublicClientApplication(
        client_id=client_id,
        authority="https://login.microsoftonline.com/common"
    )

async def ensure_valid_token() -> bool:
    """Ensure we have a valid access token, refreshing if needed."""
    global access_token, msal_app
    
    # First, try loading cached tokens
    if not access_token:
        load_tokens()
    
    # Check if current token is still valid
    if access_token and token_expires_at and time.time() < token_expires_at:
        return True
    
    # Try to refresh the token
    if not msal_app:
        msal_app = init_msal_app(get_client_id())
    
    if await refresh_access_token():
        return True
    
    # No valid token available
    access_token = None
    return False

# Global variable to store the current authentication flow
current_flow = None

@mcp.tool()
async def start_authentication() -> str:
    """
    Start the full authentication process.
    
    Returns:
        Authentication instructions with device code
    """
    global access_token, msal_app, current_flow
    
    try:
        client_id = get_client_id()
        logger.info(f"Starting authentication with client_id: {client_id[:8]}...")
        
        # Create MSAL app if not exists
        if not msal_app:
            msal_app = init_msal_app(client_id)
        
        # Start device code flow
        logger.info("Initiating device flow for authentication...")
        flow = msal_app.initiate_device_flow(scopes=SCOPES)
        
        if "user_code" not in flow:
            error_msg = flow.get('error_description', 'Unknown error in device flow')
            raise Exception(f"Failed to create device flow: {error_msg}")
        
        # Return the authentication instructions
        result = {
            "status": "authentication_required",
            "instructions": f"Go to {flow['verification_uri']} and enter code: {flow['user_code']}",
            "verification_uri": flow['verification_uri'],
            "user_code": flow['user_code'],
            "expires_in": flow.get('expires_in', 900),
            "message": "Please complete authentication, then call 'complete_authentication'"
        }
        
        # Store the flow for completion
        current_flow = flow
        
        return json.dumps(result, indent=2)
        
    except Exception as e:
        logger.error(f"Start authentication error: {str(e)}")
        return json.dumps({
            "status": "error",
            "error": str(e)
        }, indent=2)

@mcp.tool()
async def complete_authentication() -> str:
    """
    Complete the authentication process after user enters device code.
    
    Returns:
        Authentication status and user info
    """
    global access_token, msal_app, current_flow
    
    try:
        if not current_flow:
            return json.dumps({
                "status": "error",
                "error": "No authentication flow in progress. Call 'start_authentication' first."
            }, indent=2)
        
        if not msal_app:
            return json.dumps({
                "status": "error", 
                "error": "MSAL app not initialized"
            }, indent=2)
        
        logger.info("Completing device flow authentication...")
        
        # Complete the flow
        result = msal_app.acquire_token_by_device_flow(current_flow)
        
        if "access_token" in result:
            # Save tokens for future use
            save_tokens(
                result["access_token"],
                result.get("refresh_token"),
                result.get("expires_in", 3600)
            )
            
            logger.info("Authentication successful and tokens cached!")
            
            # Test the token with a basic Graph API call
            try:
                user_info = await make_graph_request("/me")
                return json.dumps({
                    "status": "success",
                    "message": "Authentication completed successfully and tokens cached for future use",
                    "user": user_info.get("displayName", "Unknown"),
                    "email": user_info.get("mail") or user_info.get("userPrincipalName", "Unknown")
                }, indent=2)
                        
            except Exception as graph_error:
                return json.dumps({
                    "status": "partial_success",
                    "message": "Got access token but Graph API test failed",
                    "graph_error": str(graph_error)
                }, indent=2)
        else:
            error_desc = result.get('error_description', 'Unknown authentication error')
            return json.dumps({
                "status": "error",
                "error": f"Authentication failed: {error_desc}"
            }, indent=2)
            
    except Exception as e:
        logger.error(f"Complete authentication error: {str(e)}")
        return json.dumps({
            "status": "error",
            "error": str(e)
        }, indent=2)
    finally:
        # Clear the flow
        current_flow = None

@mcp.tool()
async def check_authentication() -> str:
    """
    Check current authentication status and token validity.
    
    Returns:
        Authentication status information
    """
    try:
        cache_status = "enabled" if TOKEN_CACHE_ENABLED else "disabled"
        cache_file_exists = TOKEN_CACHE_FILE.exists() if TOKEN_CACHE_ENABLED else False
        
        if await ensure_valid_token():
            try:
                user_info = await make_graph_request("/me")
                time_until_expiry = int(token_expires_at - time.time()) if token_expires_at else 0
                
                return json.dumps({
                    "status": "authenticated",
                    "user": user_info.get("displayName", "Unknown"),
                    "email": user_info.get("mail") or user_info.get("userPrincipalName", "Unknown"),
                    "token_valid_for_seconds": max(0, time_until_expiry),
                    "token_valid_for_hours": round(max(0, time_until_expiry) / 3600, 1),
                    "token_caching": cache_status,
                    "cache_file_exists": cache_file_exists,
                    "cache_file_path": str(TOKEN_CACHE_FILE) if TOKEN_CACHE_ENABLED else "N/A"
                }, indent=2)
                
            except Exception as graph_error:
                return json.dumps({
                    "status": "token_invalid",
                    "error": str(graph_error),
                    "message": "Token exists but API call failed - may need re-authentication",
                    "token_caching": cache_status
                }, indent=2)
        else:
            return json.dumps({
                "status": "not_authenticated",
                "message": "No valid authentication token. Please call 'start_authentication'",
                "token_caching": cache_status,
                "cache_file_exists": cache_file_exists
            }, indent=2)
            
    except Exception as e:
        return json.dumps({
            "status": "error",
            "error": str(e),
            "token_caching": "unknown"
        }, indent=2)

async def make_graph_request(endpoint: str, method: str = "GET", data: Dict = None, use_beta: bool = False) -> Dict:
    """Make a request to Microsoft Graph API.

    Args:
        endpoint: API endpoint (e.g., "/me/onenote/notebooks")
        method: HTTP method (GET, POST, PATCH, DELETE)
        data: Request body for POST/PATCH
        use_beta: Use /beta endpoint instead of /v1.0 (needed for some operations)
    """
    # Ensure we have a valid token before making the request
    if not await ensure_valid_token():
        raise Exception("Not authenticated. Please call 'start_authentication' and 'complete_authentication' first.")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    base_url = "https://graph.microsoft.com/beta" if use_beta else GRAPH_BASE_URL
    url = f"{base_url}{endpoint}"
    
    async with httpx.AsyncClient() as client:
        if method == "GET":
            response = await client.get(url, headers=headers)
        elif method == "POST":
            response = await client.post(url, headers=headers, json=data)
        elif method == "PATCH":
            response = await client.patch(url, headers=headers, json=data)
        elif method == "DELETE":
            response = await client.delete(url, headers=headers)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")

    if response.status_code >= 400:
        raise Exception(f"Graph API error: {response.status_code} - {response.text}")

    # DELETE returns 204 No Content
    if method == "DELETE":
        return {"status": "deleted"}

    return response.json()

@mcp.tool()
async def list_notebooks() -> str:
    """
    List all OneNote notebooks.
    
    Returns:
        JSON string containing notebook information
    """
    try:
        logger.info("Making request to /me/onenote/notebooks")
        notebooks = await make_graph_request("/me/onenote/notebooks")
        logger.info(f"Graph API response received with {len(notebooks.get('value', []))} notebooks")
        
        result = []
        for notebook in notebooks.get("value", []):
            # Extract creator info
            created_by = notebook.get("createdBy", {})
            created_by_user = created_by.get("user", {})

            # Extract modifier info
            modified_by = notebook.get("lastModifiedBy", {})
            modified_by_user = modified_by.get("user", {})

            # Extract links
            links = notebook.get("links", {})

            result.append({
                "id": notebook.get("id"),
                "name": notebook.get("displayName"),
                "created": notebook.get("createdDateTime"),
                "modified": notebook.get("lastModifiedDateTime"),
                "isShared": notebook.get("isShared"),
                "userRole": notebook.get("userRole"),
                "isDefault": notebook.get("isDefault"),
                "createdBy": {
                    "name": created_by_user.get("displayName"),
                    "id": created_by_user.get("id"),
                },
                "lastModifiedBy": {
                    "name": modified_by_user.get("displayName"),
                    "id": modified_by_user.get("id"),
                },
                "webUrl": links.get("oneNoteWebUrl", {}).get("href") if links else None
            })
        
        logger.info(f"Returning {len(result)} notebooks")
        return json.dumps(result, indent=2)
    
    except Exception as e:
        logger.error(f"Error in list_notebooks: {str(e)}")
        return f"Error listing notebooks: {str(e)}"

@mcp.tool()
async def list_sections(notebook_id: str) -> str:
    """
    List sections in a specific notebook.
    
    Args:
        notebook_id: ID of the notebook to list sections from
    
    Returns:
        JSON string containing section information
    """
    try:
        sections = await make_graph_request(f"/me/onenote/notebooks/{notebook_id}/sections")
        
        result = []
        for section in sections.get("value", []):
            result.append({
                "id": section.get("id"),
                "name": section.get("displayName"),
                "created": section.get("createdDateTime"),
                "modified": section.get("lastModifiedDateTime")
            })
        
        return json.dumps(result, indent=2)
    
    except Exception as e:
        return f"Error listing sections: {str(e)}"

# =============================================================================
# Section Groups Tools
# =============================================================================

@mcp.tool()
async def list_section_groups(notebook_id: str) -> str:
    """
    List all section groups in a specific notebook.

    Args:
        notebook_id: ID of the notebook to list section groups from

    Returns:
        JSON string containing section group information
    """
    try:
        section_groups = await make_graph_request(
            f"/me/onenote/notebooks/{notebook_id}/sectionGroups"
        )

        result = []
        for group in section_groups.get("value", []):
            # Extract creator info
            created_by = group.get("createdBy", {})
            created_by_user = created_by.get("user", {})

            # Extract modifier info
            modified_by = group.get("lastModifiedBy", {})
            modified_by_user = modified_by.get("user", {})

            result.append({
                "id": group.get("id"),
                "name": group.get("displayName"),
                "created": group.get("createdDateTime"),
                "modified": group.get("lastModifiedDateTime"),
                "sectionsUrl": group.get("sectionsUrl"),
                "sectionGroupsUrl": group.get("sectionGroupsUrl"),
                "self": group.get("self"),
                "createdBy": {
                    "name": created_by_user.get("displayName"),
                    "id": created_by_user.get("id"),
                },
                "lastModifiedBy": {
                    "name": modified_by_user.get("displayName"),
                    "id": modified_by_user.get("id"),
                },
            })

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error listing section groups: {str(e)}"


@mcp.tool()
async def get_section_group(section_group_id: str) -> str:
    """
    Get details of a specific section group.

    Args:
        section_group_id: ID of the section group to retrieve

    Returns:
        JSON string containing section group details
    """
    try:
        group = await make_graph_request(
            f"/me/onenote/sectionGroups/{section_group_id}"
        )

        # Extract creator info
        created_by = group.get("createdBy", {})
        created_by_user = created_by.get("user", {})

        # Extract modifier info
        modified_by = group.get("lastModifiedBy", {})
        modified_by_user = modified_by.get("user", {})

        result = {
            "id": group.get("id"),
            "name": group.get("displayName"),
            "created": group.get("createdDateTime"),
            "modified": group.get("lastModifiedDateTime"),
            "sectionsUrl": group.get("sectionsUrl"),
            "sectionGroupsUrl": group.get("sectionGroupsUrl"),
            "self": group.get("self"),
            "createdBy": {
                "name": created_by_user.get("displayName"),
                "id": created_by_user.get("id"),
            },
            "lastModifiedBy": {
                "name": modified_by_user.get("displayName"),
                "id": modified_by_user.get("id"),
            },
        }

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error getting section group: {str(e)}"


@mcp.tool()
async def list_sections_in_group(section_group_id: str) -> str:
    """
    List all sections within a specific section group.

    Args:
        section_group_id: ID of the section group to list sections from

    Returns:
        JSON string containing section information
    """
    try:
        sections = await make_graph_request(
            f"/me/onenote/sectionGroups/{section_group_id}/sections"
        )

        result = []
        for section in sections.get("value", []):
            result.append({
                "id": section.get("id"),
                "name": section.get("displayName"),
                "created": section.get("createdDateTime"),
                "modified": section.get("lastModifiedDateTime"),
                "isDefault": section.get("isDefault"),
                "pagesUrl": section.get("pagesUrl"),
            })

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error listing sections in group: {str(e)}"


@mcp.tool()
async def list_nested_section_groups(section_group_id: str) -> str:
    """
    List all nested section groups within a specific section group.
    OneNote supports hierarchical section groups (groups within groups).

    Args:
        section_group_id: ID of the parent section group

    Returns:
        JSON string containing nested section group information
    """
    try:
        section_groups = await make_graph_request(
            f"/me/onenote/sectionGroups/{section_group_id}/sectionGroups"
        )

        result = []
        for group in section_groups.get("value", []):
            # Extract creator info
            created_by = group.get("createdBy", {})
            created_by_user = created_by.get("user", {})

            # Extract modifier info
            modified_by = group.get("lastModifiedBy", {})
            modified_by_user = modified_by.get("user", {})

            result.append({
                "id": group.get("id"),
                "name": group.get("displayName"),
                "created": group.get("createdDateTime"),
                "modified": group.get("lastModifiedDateTime"),
                "sectionsUrl": group.get("sectionsUrl"),
                "sectionGroupsUrl": group.get("sectionGroupsUrl"),
                "self": group.get("self"),
                "createdBy": {
                    "name": created_by_user.get("displayName"),
                    "id": created_by_user.get("id"),
                },
                "lastModifiedBy": {
                    "name": modified_by_user.get("displayName"),
                    "id": modified_by_user.get("id"),
                },
            })

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error listing nested section groups: {str(e)}"


@mcp.tool()
async def create_section_group(notebook_id: str, name: str) -> str:
    """
    Create a new section group in a notebook.

    Args:
        notebook_id: ID of the notebook to create the section group in
        name: Name of the new section group (max 50 chars, no special chars: ?*\\/:<>|&#''%~)

    Returns:
        JSON string with the created section group information
    """
    try:
        data = {"displayName": name}

        group = await make_graph_request(
            f"/me/onenote/notebooks/{notebook_id}/sectionGroups",
            method="POST",
            data=data
        )

        result = {
            "status": "success",
            "message": f"Section group '{name}' created successfully",
            "sectionGroup": {
                "id": group.get("id"),
                "name": group.get("displayName"),
                "created": group.get("createdDateTime"),
                "sectionsUrl": group.get("sectionsUrl"),
                "sectionGroupsUrl": group.get("sectionGroupsUrl"),
            }
        }

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error creating section group: {str(e)}"


@mcp.tool()
async def create_section_in_group(section_group_id: str, name: str) -> str:
    """
    Create a new section within a section group.

    Args:
        section_group_id: ID of the section group to create the section in
        name: Name of the new section

    Returns:
        JSON string with the created section information
    """
    try:
        data = {"displayName": name}

        section = await make_graph_request(
            f"/me/onenote/sectionGroups/{section_group_id}/sections",
            method="POST",
            data=data
        )

        result = {
            "status": "success",
            "message": f"Section '{name}' created successfully in section group",
            "section": {
                "id": section.get("id"),
                "name": section.get("displayName"),
                "created": section.get("createdDateTime"),
                "pagesUrl": section.get("pagesUrl"),
            }
        }

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error creating section in group: {str(e)}"


@mcp.tool()
async def create_nested_section_group(parent_section_group_id: str, name: str) -> str:
    """
    Create a nested section group within an existing section group.
    OneNote supports hierarchical section groups (groups within groups).

    Args:
        parent_section_group_id: ID of the parent section group
        name: Name of the new nested section group (max 50 chars)

    Returns:
        JSON string with the created nested section group information
    """
    try:
        data = {"displayName": name}

        group = await make_graph_request(
            f"/me/onenote/sectionGroups/{parent_section_group_id}/sectionGroups",
            method="POST",
            data=data
        )

        result = {
            "status": "success",
            "message": f"Nested section group '{name}' created successfully",
            "sectionGroup": {
                "id": group.get("id"),
                "name": group.get("displayName"),
                "created": group.get("createdDateTime"),
                "sectionsUrl": group.get("sectionsUrl"),
                "sectionGroupsUrl": group.get("sectionGroupsUrl"),
            }
        }

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error creating nested section group: {str(e)}"


# =============================================================================
# Page Tools
# =============================================================================

@mcp.tool()
async def list_pages(section_id: str) -> str:
    """
    List pages in a specific section, ordered by display position.

    Args:
        section_id: ID of the section to list pages from

    Returns:
        JSON string containing page information with hierarchy (level, order)
    """
    try:
        # Explicitly request level and order properties with $select
        # Order by 'order' to get pages in display order (as shown in OneNote UI)
        endpoint = (
            f"/me/onenote/sections/{section_id}/pages"
            "?$select=id,title,createdDateTime,lastModifiedDateTime,contentUrl,level,order"
            "&$orderby=order"
        )
        pages = await make_graph_request(endpoint)

        result = []
        for page in pages.get("value", []):
            result.append({
                "id": page.get("id"),
                "title": page.get("title"),
                "created": page.get("createdDateTime"),
                "modified": page.get("lastModifiedDateTime"),
                "content_url": page.get("contentUrl"),
                "level": page.get("level", 0),  # 0 = top-level, 1+ = subpage
                "order": page.get("order")  # Display order within section
            })

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error listing pages: {str(e)}"

@mcp.tool()
async def get_page_content(page_id: str) -> str:
    """
    Get the content of a specific page.
    
    Args:
        page_id: ID of the page to retrieve content from
    
    Returns:
        Page content as HTML or error message
    """
    try:
        # Get page content (returns HTML)
        async with httpx.AsyncClient() as client:
            headers = {"Authorization": f"Bearer {access_token}"}
            response = await client.get(
                f"{GRAPH_BASE_URL}/me/onenote/pages/{page_id}/content",
                headers=headers
            )
            
            if response.status_code >= 400:
                return f"Error getting page content: {response.status_code} - {response.text}"
            
            return response.text
    
    except Exception as e:
        return f"Error getting page content: {str(e)}"

@mcp.tool()
async def clear_token_cache() -> str:
    """
    Clear the stored authentication tokens.
    
    Returns:
        Status message
    """
    global access_token, refresh_token, token_expires_at
    
    try:
        # Clear in-memory tokens
        access_token = None
        refresh_token = None
        token_expires_at = None
        
        # Remove cache file
        if TOKEN_CACHE_FILE.exists():
            TOKEN_CACHE_FILE.unlink()
            
        return json.dumps({
            "status": "success",
            "message": "Token cache cleared. You will need to re-authenticate."
        }, indent=2)
        
    except Exception as e:
        return json.dumps({
            "status": "error",
            "error": str(e)
        }, indent=2)

@mcp.tool()
async def create_notebook(name: str, description: str = None) -> str:
    """
    Create a new OneNote notebook.
    
    Args:
        name: Name of the new notebook
        description: Optional description for the notebook
    
    Returns:
        JSON string with the created notebook information
    """
    try:
        data = {"displayName": name}
        if description:
            data["description"] = description
            
        notebook = await make_graph_request("/me/onenote/notebooks", method="POST", data=data)
        
        result = {
            "status": "success",
            "message": f"Notebook '{name}' created successfully",
            "notebook": {
                "id": notebook.get("id"),
                "name": notebook.get("displayName"),
                "created": notebook.get("createdDateTime")
            }
        }
        
        return json.dumps(result, indent=2)
    
    except Exception as e:
        return f"Error creating notebook: {str(e)}"

@mcp.tool()
async def create_section(notebook_id: str, name: str) -> str:
    """
    Create a new section in a OneNote notebook.
    
    Args:
        notebook_id: ID of the notebook to create the section in
        name: Name of the new section
    
    Returns:
        JSON string with the created section information
    """
    try:
        data = {"displayName": name}
        
        section = await make_graph_request(
            f"/me/onenote/notebooks/{notebook_id}/sections", 
            method="POST", 
            data=data
        )
        
        result = {
            "status": "success",
            "message": f"Section '{name}' created successfully",
            "section": {
                "id": section.get("id"),
                "name": section.get("displayName"),
                "created": section.get("createdDateTime")
            }
        }
        
        return json.dumps(result, indent=2)
    
    except Exception as e:
        return f"Error creating section: {str(e)}"

@mcp.tool()
async def create_page(section_id: str, title: str, content_html: str = None) -> str:
    """
    Create a new page in a OneNote section.
    
    Args:
        section_id: ID of the section to create the page in
        title: Title of the new page
        content_html: Optional HTML content for the page body
    
    Returns:
        JSON string with the created page information
    """
    try:
        # Build the HTML structure for the page
        if content_html:
            # Ensure content is wrapped in proper OneNote HTML structure
            if not content_html.strip().startswith('<html>'):
                page_html = f"""
<!DOCTYPE html>
<html>
<head>
    <title>{title}</title>
    <meta name="created" content="{time.strftime('%Y-%m-%dT%H:%M:%S.0000000')}" />
</head>
<body>
    <div>
        <h1>{title}</h1>
        <div>{content_html}</div>
    </div>
</body>
</html>"""
            else:
                page_html = content_html
        else:
            # Create a basic page with just the title
            page_html = f"""
<!DOCTYPE html>
<html>
<head>
    <title>{title}</title>
    <meta name="created" content="{time.strftime('%Y-%m-%dT%H:%M:%S.0000000')}" />
</head>
<body>
    <div>
        <h1>{title}</h1>
        <p>Page created by OneNote MCP Server</p>
    </div>
</body>
</html>"""
        
        # OneNote API expects multipart form data for page creation
        async with httpx.AsyncClient() as client:
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/xhtml+xml"
            }
            
            response = await client.post(
                f"{GRAPH_BASE_URL}/me/onenote/sections/{section_id}/pages",
                headers=headers,
                content=page_html
            )
            
            if response.status_code >= 400:
                return f"Error creating page: {response.status_code} - {response.text}"
            
            page = response.json()
        
        result = {
            "status": "success",
            "message": f"Page '{title}' created successfully",
            "page": {
                "id": page.get("id"),
                "title": page.get("title"),
                "created": page.get("createdDateTime"),
                "content_url": page.get("contentUrl")
            }
        }
        
        return json.dumps(result, indent=2)
    
    except Exception as e:
        return f"Error creating page: {str(e)}"

@mcp.tool()
async def update_page_content(page_id: str, content_html: str, target_element: str = "body") -> str:
    """
    Update the content of an existing OneNote page.
    
    Args:
        page_id: ID of the page to update
        content_html: New HTML content to add/replace
        target_element: Target element to update (default: "body")
    
    Returns:
        Status message
    """
    try:
        # OneNote PATCH API for updating page content
        patch_data = [
            {
                "target": target_element,
                "action": "append",
                "content": content_html
            }
        ]
        
        async with httpx.AsyncClient() as client:
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json"
            }
            
            response = await client.patch(
                f"{GRAPH_BASE_URL}/me/onenote/pages/{page_id}/content",
                headers=headers,
                json=patch_data
            )
            
            if response.status_code >= 400:
                return f"Error updating page: {response.status_code} - {response.text}"
        
        result = {
            "status": "success",
            "message": "Page content updated successfully",
            "page_id": page_id
        }
        
        return json.dumps(result, indent=2)
    
    except Exception as e:
        return f"Error updating page content: {str(e)}"


# =============================================================================
# Delete Tools
# =============================================================================

@mcp.tool()
async def delete_page(page_id: str) -> str:
    """
    Delete a page from OneNote.
    WARNING: This action is irreversible. The page will be permanently deleted.

    Note: Only page deletion is supported by Microsoft Graph API.
    Sections, section groups, and notebooks cannot be deleted via the API.

    Args:
        page_id: ID of the page to delete

    Returns:
        JSON string with deletion status
    """
    try:
        await make_graph_request(
            f"/me/onenote/pages/{page_id}",
            method="DELETE"
        )

        result = {
            "status": "success",
            "message": "Page deleted successfully",
            "page_id": page_id
        }

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error deleting page: {str(e)}"


def _strip_onenote_id_prefix(onenote_id: str) -> str:
    """
    Strip the '0-' prefix from OneNote IDs for use with OneDrive API.

    OneNote IDs from Graph API have format: 0-D5F846AE8B2C1F44!s...
    OneDrive API needs the format: D5F846AE8B2C1F44!s...
    """
    if onenote_id.startswith("0-"):
        return onenote_id[2:]
    return onenote_id


@mcp.tool()
async def delete_section(section_id: str) -> str:
    """
    Delete a section from OneNote via OneDrive API.
    WARNING: This action is irreversible. The section and all its pages will be permanently deleted.

    Note: Microsoft Graph OneNote API does not support section deletion directly.
    This uses the OneDrive API workaround since OneNote sections are stored as .one files in OneDrive.
    Requires Files.ReadWrite scope.

    Args:
        section_id: ID of the section to delete (OneNote ID format accepted)

    Returns:
        JSON string with deletion status
    """
    try:
        # Convert OneNote ID to OneDrive ID (strip "0-" prefix)
        drive_item_id = _strip_onenote_id_prefix(section_id)

        # Use OneDrive API to delete the section (stored as a .one file)
        await make_graph_request(
            f"/me/drive/items/{drive_item_id}",
            method="DELETE"
        )

        result = {
            "status": "success",
            "message": "Section deleted successfully via OneDrive API",
            "section_id": section_id,
            "note": "Section and all its pages have been permanently deleted"
        }

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error deleting section: {str(e)}"


@mcp.tool()
async def delete_section_group(section_group_id: str) -> str:
    """
    Delete a section group from OneNote via OneDrive API.
    WARNING: This action is irreversible. The section group and all its contents will be permanently deleted.

    Note: Microsoft Graph OneNote API does not support section group deletion directly.
    This uses the OneDrive API workaround since OneNote section groups are stored as folders in OneDrive.
    Requires Files.ReadWrite scope.

    Args:
        section_group_id: ID of the section group to delete (OneNote ID format accepted)

    Returns:
        JSON string with deletion status
    """
    try:
        # Convert OneNote ID to OneDrive ID (strip "0-" prefix)
        drive_item_id = _strip_onenote_id_prefix(section_group_id)

        # Use OneDrive API to delete the section group (stored as a folder)
        await make_graph_request(
            f"/me/drive/items/{drive_item_id}",
            method="DELETE"
        )

        result = {
            "status": "success",
            "message": "Section group deleted successfully via OneDrive API",
            "section_group_id": section_group_id,
            "note": "Section group and all its contents have been permanently deleted"
        }

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error deleting section group: {str(e)}"


# =============================================================================
# Copy Tools
# =============================================================================

@mcp.tool()
async def copy_page_to_section(page_id: str, target_section_id: str) -> str:
    """
    Copy a page to another section.
    The original page remains in its current location.

    Note: Uses the /beta endpoint as /v1.0 returns 501 for this operation.

    Args:
        page_id: ID of the page to copy
        target_section_id: ID of the destination section

    Returns:
        JSON string with copy operation status
    """
    try:
        data = {"id": target_section_id}

        # Note: copyToSection requires /beta endpoint (/v1.0 returns 501)
        # This is an async operation that returns an operation URL
        result_data = await make_graph_request(
            f"/me/onenote/pages/{page_id}/copyToSection",
            method="POST",
            data=data,
            use_beta=True
        )

        result = {
            "status": "success",
            "message": "Page copy initiated successfully",
            "page_id": page_id,
            "target_section_id": target_section_id,
            "operation": result_data
        }

        return json.dumps(result, indent=2)

    except Exception as e:
        return f"Error copying page: {str(e)}"


def main():
    """Main entry point for the server."""
    # Log token caching configuration
    cache_status = "enabled" if TOKEN_CACHE_ENABLED else "disabled"
    logger.info(f"OneNote MCP Server starting - Token caching: {cache_status}")
    
    if TOKEN_CACHE_ENABLED:
        logger.info(f"Token cache file: {TOKEN_CACHE_FILE}")
        # Try to load cached tokens on startup
        if load_tokens():
            logger.info("Cached tokens loaded successfully")
        else:
            logger.info("No valid cached tokens found")
    
    mcp.run()

if __name__ == "__main__":
    main()
