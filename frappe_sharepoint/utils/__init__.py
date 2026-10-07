from __future__ import absolute_import

import frappe
from frappe import _
import requests

# Get access token using client credentials flow
def get_access_token(tenant_id, client_id, client_secret):
    """
    Authenticate with Azure AD using client credentials flow
    Returns access token for Microsoft Graph API
    """
    frappe.logger().info(f"[Azure Auth] Starting authentication for tenant: {tenant_id[:8]}...")
    frappe.logger().info(f"[Azure Auth] Client ID: {client_id[:8]}...")
    
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    frappe.logger().info(f"[Azure Auth] Token URL: {token_url}")
    
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    
    try:
        frappe.logger().info(f"[Azure Auth] Sending authentication request...")
        response = requests.post(token_url, data=data, timeout=30)
        frappe.logger().info(f"[Azure Auth] Response status: {response.status_code}")
        
        if response.ok:
            token = response.json().get('access_token')
            if token:
                frappe.logger().info(f"[Azure Auth] Successfully obtained access token (length: {len(token)})")
                return token
            else:
                frappe.logger().error(f"[Azure Auth] No access_token in response: {response.json()}")
                frappe.log_error("Azure AD Token Error", "No access_token in response")
                return None
        else:
            frappe.logger().error(f"[Azure Auth] Authentication failed with status {response.status_code}")
            frappe.logger().error(f"[Azure Auth] Response: {response.text}")
            frappe.log_error("Azure AD Token Error", f"Status: {response.status_code}, Response: {response.text}")
            return None
            
    except requests.exceptions.Timeout as e:
        frappe.logger().error(f"[Azure Auth] Request timeout: {str(e)}")
        frappe.log_error("Azure AD Authentication Timeout", str(e))
        return None
    except requests.exceptions.ConnectionError as e:
        frappe.logger().error(f"[Azure Auth] Connection error: {str(e)}")
        frappe.log_error("Azure AD Connection Error", str(e))
        return None
    except requests.exceptions.RequestException as e:
        frappe.logger().error(f"[Azure Auth] Request exception: {str(e)}")
        frappe.log_error("Azure AD Request Error", str(e))
        return None
    except Exception as e:
        frappe.logger().error(f"[Azure Auth] Unexpected error: {str(e)}")
        frappe.log_error("Azure AD Authentication Error", str(e))
        return None

# Make request headers with bearer token
def get_request_header(settings):
    """
    Generate authorization headers using Azure AD credentials
    """
    frappe.logger().info(f"[Request Header] Generating authorization headers")
    
    try:
        # Validate settings
        if not settings.tenant_id:
            frappe.logger().error(f"[Request Header] Missing tenant_id in settings")
            frappe.throw(_("Tenant ID is not configured in SharePoint Settings"))
        
        if not settings.client_id:
            frappe.logger().error(f"[Request Header] Missing client_id in settings")
            frappe.throw(_("Client ID is not configured in SharePoint Settings"))
        
        client_secret = settings.get_password("client_secret")
        if not client_secret:
            frappe.logger().error(f"[Request Header] Missing client_secret in settings")
            frappe.throw(_("Client Secret is not configured in SharePoint Settings"))
        
        frappe.logger().info(f"[Request Header] Settings validated, requesting access token")
        
        access_token = get_access_token(
            settings.tenant_id,
            settings.client_id,
            client_secret
        )
        
        if not access_token:
            frappe.logger().error(f"[Request Header] Failed to obtain access token")
            frappe.throw(_("Failed to authenticate with Azure AD. Please check your credentials in SharePoint Settings."))
        
        frappe.logger().info(f"[Request Header] Successfully generated authorization header")
        headers = {'Authorization': f'Bearer {access_token}'}
        return headers
        
    except Exception as e:
        frappe.logger().error(f"[Request Header] Exception: {str(e)}")
        raise
    
# General API request handler
def make_request(request, url, headers, body=None):
    """
    Make HTTP requests to Microsoft Graph API with comprehensive error handling
    """
    frappe.logger().info(f"[API Request] Method: {request}, URL: {url[:100]}...")
    frappe.logger().info(f"[API Request] Headers present: {list(headers.keys())}")
    
    # Set timeout for all requests (30 seconds)
    timeout = 30
    
    try:
        if request == 'POST':
            frappe.logger().info(f"[API Request] Making POST request with JSON body")
            response = requests.post(url, headers=headers, json=body, timeout=timeout)
        elif request == 'PATCH':
            frappe.logger().info(f"[API Request] Making PATCH request with JSON body")
            response = requests.patch(url, headers=headers, json=body, timeout=timeout)
        elif request == 'GET':
            frappe.logger().info(f"[API Request] Making GET request")
            response = requests.get(url, headers=headers, timeout=timeout)
        elif request == 'DELETE':
            frappe.logger().info(f"[API Request] Making DELETE request")
            response = requests.delete(url, headers=headers, timeout=timeout)
        elif request == "PUT":
            frappe.logger().info("[API Request] Making PUT request with binary body")
            response = requests.put(url, headers=headers, data=body, timeout=timeout)
        else:
            frappe.logger().error(f"[API Request] Unsupported request method: {request}")
            frappe.log_error("Unsupported HTTP Method", f"Method: {request}")
            return None
        
        frappe.logger().info(f"[API Request] Response status: {response.status_code}")
        
        # Log response details for non-200 responses
        if not response.ok:
            frappe.logger().warning(f"[API Request] Non-OK response: {response.status_code}")
            frappe.logger().warning(f"[API Request] Response headers: {dict(response.headers)}")
            try:
                # Try to get JSON error details
                error_data = response.json()
                frappe.logger().error(f"[API Request] Error response: {error_data}")
            except:
                # If not JSON, log text
                frappe.logger().error(f"[API Request] Error text: {response.text[:500]}")
        
        return response
        
    except requests.exceptions.Timeout as e:
        frappe.logger().error(f"[API Request] Timeout after {timeout}s: {str(e)}")
        frappe.log_error("Microsoft Graph API Timeout", f"URL: {url}\nError: {str(e)}")
        # Return a mock response object with error details
        return create_error_response(f"Request timeout after {timeout} seconds", 408)
        
    except requests.exceptions.ConnectionError as e:
        frappe.logger().error(f"[API Request] Connection error: {str(e)}")
        frappe.log_error("Microsoft Graph API Connection Error", f"URL: {url}\nError: {str(e)}")
        return create_error_response(f"Connection error: {str(e)}", 503)
        
    except requests.exceptions.HTTPError as e:
        frappe.logger().error(f"[API Request] HTTP error: {str(e)}")
        frappe.log_error("Microsoft Graph API HTTP Error", f"URL: {url}\nError: {str(e)}")
        return create_error_response(f"HTTP error: {str(e)}", 500)
        
    except requests.exceptions.RequestException as e:
        frappe.logger().error(f"[API Request] Request exception: {str(e)}")
        frappe.log_error("Microsoft Graph API Request Error", f"URL: {url}\nError: {str(e)}")
        return create_error_response(f"Request error: {str(e)}", 500)
        
    except Exception as e:
        frappe.logger().error(f"[API Request] Unexpected error: {str(e)}")
        frappe.log_error("Microsoft Graph API Unexpected Error", f"URL: {url}\nError: {str(e)}")
        return create_error_response(f"Unexpected error: {str(e)}", 500)


# Helper to create error response objects
def create_error_response(error_message, status_code):
    """
    Create a mock response object for error cases
    This ensures we always return a response-like object with .ok and .text attributes
    """
    class ErrorResponse:
        def __init__(self, message, code):
            self.text = message
            self.status_code = code
            self.ok = False
            self.content = message.encode('utf-8')
            
        def json(self):
            return {"error": {"message": self.text}}
    
    return ErrorResponse(error_message, status_code)