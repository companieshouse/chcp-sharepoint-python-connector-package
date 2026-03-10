    
"""
sharepoint_client.py

This module provides the `SharePointClient` class for interacting with Microsoft 
SharePoint via the Microsoft Graph API.

Features:
- Authenticate using Azure AD credentials (tenant ID, client ID, client secret)
- Retrieve site, drive and file IDs
- List Sharepoint folder contents
- Download files from SharePoint
- Load Sharepoint files into memory
- Upload files to SharePoint

Raises:
    Exception: If authentication or file operations fail
    ValueError: For invalid URLs or parameters
"""

import msal
import requests
import os
from pathlib import Path
from urllib.parse import urlparse, unquote


class SharePointClient:
     
    def __init__(self, tenant_id: str, client_id: str, client_secret: str, scopes: list[str]):
        self.tenant_id: str = tenant_id
        self.client_id: str = client_id
        self.client_secret: str = client_secret
        self.authority: str = f"https://login.microsoftonline.com/{tenant_id}"
        self.scopes: list[str] = scopes
        self.ms_graph_url: str = "https://graph.microsoft.com/v1.0"
        self.app: msal.ConfidentialClientApplication = msal.ConfidentialClientApplication(
            client_id,
            authority=self.authority,
            client_credential=client_secret
        )
        self.access_token = self.get_access_token()  # Initialize and store the access token upon instantiation

    def get_access_token(self):
        result = self.app.acquire_token_for_client(scopes=self.scopes)
        if "access_token" not in result:
            raise Exception(f"Token acquisition failed: {result.get('error_description')}")
        print("Access token created")
        return result["access_token"]

    def get_site_id(self, site_url: str) -> str:
        full_url = f'{self.ms_graph_url}/sites/{site_url}'
        response = requests.get(full_url, headers={'Authorization': f'Bearer {self.access_token}'})
        return response.json().get('id')
    
    def get_drives(self, site_id: str) -> dict:
        drives_url = f'{self.ms_graph_url}/sites/{site_id}/drives'
        response = requests.get(drives_url, headers={'Authorization': f'Bearer {self.access_token}'})
        drives = response.json().get('value', [])
        drives_dict = {}
        for drive in drives:
            drives_dict[drive['id']] = drive['name']
        return drives_dict
    
    def resolve_drive_id(self, site_id: str, drive_name: str) -> str:
        """
        Given a site_id and drive name, return the drive ID.
        Raises ValueError if the drive is not found.
        """
        drives = self.get_drives(site_id)
       
        for id_, name in drives.items():
            if name == drive_name:
                return id_
        raise ValueError(f"Drive '{drive_name}' not found for site '{site_id}'.")
    
    def get_folder_content(
            self,
            site_id: str,
            drive_id: str,
            folder_path: str = 'root'
        ) -> list:
            """
            Returns a list of all items (dicts with 'id' and 'name') in the specified 
            folder, handling pagination.
            """
            if folder_path == 'root' or folder_path.strip() == '':
                folder_url = f'{self.ms_graph_url}/sites/{site_id}/drives/{drive_id}/root/children'
            else:
                folder_path_clean = folder_path.strip('/\\')
                folder_url = f'{self.ms_graph_url}/sites/{site_id}/drives/{drive_id}/root:/{folder_path_clean}:/children'

            headers = {'Authorization': f'Bearer {self.access_token}'}
            next_url = folder_url
            folder_dict = {}

            while next_url:
                response = requests.get(next_url, headers=headers)
                data = response.json()
                items_data = data.get('value', [])
                for item in items_data:
                    folder_dict[item['id']] = item['name']
                next_url = data.get('@odata.nextLink')
            return folder_dict
    
    def resolve_file_id(
            self, 
            site_id: str, 
            drive_id: str, 
            file_name: str, 
            folder_path: str = 'root'
        ) -> str:
            """
            Given a file name and folder, return the file ID.
            Raises ValueError if the file is not found.
            """
            files = self.get_folder_content(site_id, drive_id, folder_path)
    
            # try to match by name
            for id_, name in files.items():
                if name == file_name:
                    return id_
            raise ValueError(
                f"File '{file_name}' not found in folder "
                f"'{folder_path}' for drive '{drive_id}'."
            )

    def get_file_name_from_id(self, site_id: str, drive_id: str, file_id: str) -> str:
        """
        Fetch the file name from Microsoft Graph API using file_id.
        Returns the file name as a string.
        Raises Exception if the file name cannot be determined.
        """
        headers = {'Authorization': f'Bearer {self.access_token}'}
        metadata_url = f"{self.ms_graph_url}/sites/{site_id}/drives/{drive_id}/items/{file_id}"
        metadata_response = requests.get(metadata_url, headers=headers)
        if metadata_response.status_code != 200:
            raise Exception(
                "Failed to fetch file metadata: "
                f"{metadata_response.status_code} - {metadata_response.reason}"
            )
        file_name = metadata_response.json().get('name')
        if not file_name:
            raise Exception("Could not determine file name from metadata.")
        return file_name
    
    @staticmethod
    def find_spaced_drive_name(
            no_space_drive_name: str, 
            spaced_drive_name_list: list
        ) -> str:
        """
        Sharepoint removes the spaces in drive names in the url. This function takes a 
        drive name without spaces as input, and finds the full drive name if it exists
        in the specified drive names list (that contain spaces).
        """
        for s in spaced_drive_name_list:
            if s.replace(" ", "") == no_space_drive_name:
                return s  # Return the original string with spaces
        raise ValueError("Drive name not found in the list.")

    def parse_url_to_ids(self, url: str) -> dict:
        """
        Parse a SharePoint file URL and return site_id, drive_id, and file_id as a dictionary.
        Example URL:
        https://yourtenant.sharepoint.com/sites/YourSite/YourDrive/YourFile.csv
        """
        
        parsed = urlparse(url)
        path_parts = [unquote(part) for part in parsed.path.strip('/').split('/')]
        # Find site name
        if 'sites' in path_parts:
            site_index = path_parts.index('sites')
            site_name = path_parts[site_index + 1]
        else:
            raise ValueError("URL does not contain a site name.")
        
        drive_name_no_spaces = path_parts[site_index + 2] if len(path_parts) > site_index + 2 else ''
        folder_path = '/'.join(path_parts[site_index + 3:-1]) if len(path_parts) > site_index + 2 else ''
        file_name = path_parts[-1]

        site_url = f"{parsed.netloc}:/sites/{site_name}"
        site_id = self.get_site_id(site_url)
        drives = self.get_drives(site_id)
        drive_name = self.find_spaced_drive_name(drive_name_no_spaces, drives.values())
        
       
        drive_id = self.resolve_drive_id(site_id, drive_name)
        file_id = self.resolve_file_id(site_id, drive_id, file_name, folder_path)
      
        return {"site_id": site_id, "drive_id": drive_id, "file_id": file_id}
    
    def download_file_to_disk(
            self, 
            site_id: str, 
            drive_id: str, 
            file_id: str,
            local_path: str = 'downloads'
        ) -> None:
        headers = {'Authorization': f'Bearer {self.access_token}'}

        # get file name
        file_name = self.get_file_name_from_id(site_id, drive_id, file_id)

        # Build download URL
        download_url = (
            f"{self.ms_graph_url}/sites/{site_id}/drives/"
            f"{drive_id}/items/{file_id}/content"
        )
        response = requests.get(download_url, headers=headers)
        if response.status_code == 200:
            full_path = os.path.join(local_path, file_name)
            with open(full_path, 'wb') as file:
                file.write(response.content)
            print(f"File downloaded: {full_path}")
        else:
            print(
                f"Failed to download {file_name}:" 
                f"{response.status_code} - {response.reason}"
            )

    def download_file_bytes(self, site_id: str, drive_id: str, file_id: str) -> bytes:
        """
        Download a file from SharePoint and return its content as bytes in memory.
        This can be used to load data directly into pandas or other libraries without 
        saving to disk.
        """
        download_url = (
            f"{self.ms_graph_url}/sites/{site_id}/drives/"
            f"{drive_id}/items/{file_id}/content"
        )
        headers = {'Authorization': f'Bearer {self.access_token}'}
        response = requests.get(download_url, headers=headers)
        if response.status_code == 200:
            return response.content
        else:
            raise Exception(
                f"Failed to download file as bytes: {response.status_code} - {response.reason}"
            )

    def upload_small_file(
            self, 
            local_file_path: str | Path, 
            site_id: str, 
            drive_id: str, 
            upload_path: str | Path):
        """Uploads a file <4MB to SharePoint."""
        
        with open(local_file_path, "rb") as file:
            file_bytes = file.read()
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}"
            f"/drives/{drive_id}/root:/{upload_path}:/content"
        )
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/octet-stream",
        }
        r = requests.put(url, headers=headers, data=file_bytes)
        r.raise_for_status()
        return r.json()

    def create_upload_session(self, site_id, drive_id, upload_path):
        """Create an upload session for use by upload_large_file()"""
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}"
            f"/root:/{upload_path}:/createUploadSession"
        )
        headers = {"Authorization": f"Bearer {self.access_token}"}
        r = requests.post(url, headers=headers)
        r.raise_for_status()
        return r.json()["uploadUrl"]

    def upload_large_file(upload_url, local_file_path):
        """Uploads a file >4MB to SharePoint."""
        CHUNK_SIZE = 5 * 1024 * 1024
        file_size = os.path.getsize(local_file_path)
        bytes_sent = 0
        with open(local_file_path, "rb") as f:
            while bytes_sent < file_size:
                chunk = f.read(CHUNK_SIZE)
                chunk_size = len(chunk)
                start = bytes_sent
                end = bytes_sent + chunk_size - 1
                headers = {
                    "Content-Length": str(chunk_size),
                    "Content-Range": f"bytes {start}-{end}/{file_size}",
                }
                r = requests.put(upload_url, headers=headers, data=chunk)
                r.raise_for_status()
                if r.status_code in (200, 201):
                    # return response when all chunks have been successfully uploaded
                    return r.json()
                bytes_sent = end + 1
                print(f"Sent {bytes_sent}/{file_size} bytes ({(bytes_sent/file_size)*100:.2f}%)")
        raise Exception("Upload did not complete properly.")

    def upload_file(self, local_file_path, site_id, drive_id, upload_folder_path=None):
        """
        Uploads a file to SharePoint. 
        Uses simple upload for files <= chunk_threshold_mb MB, 
        otherwise uses upload session (chunked).
        """
        local_file_path = Path(local_file_path) # convert to a path if a string is supplied
        filename = local_file_path.name

        if upload_folder_path:
            upload_folder_path = Path(upload_folder_path)
            upload_path = upload_folder_path / filename
        else:
            upload_path = filename

        file_size = os.path.getsize(local_file_path)

        # Sharepoint requires files above 4MB to use an upload session. We define this 
        # threshold here (and keep it to 3MB for safety)
        THRESHOLD = 3 * 1024 * 1024

        if file_size <= THRESHOLD:
            result = self.upload_small_file(local_file_path, site_id, drive_id, upload_path)
        else:
            upload_url = self.create_upload_session(site_id, drive_id, upload_path)
            result = self.upload_large_file(upload_url, local_file_path)
        print("\nUpload Complete")
        print("SharePoint URL:", result.get("webUrl"))
        return result

    @staticmethod
    def to_graph_site_url(sharepoint_url: str) -> str:
        """
        Convert a full SharePoint URL into the Microsoft Graph API site identifier format.
        
        Example:
            Input:  https://contoso.sharepoint.com/sites/Marketing/SitePages/Home.aspx
            Output: contoso.sharepoint.com:/sites/Marketing
        """
        
        parsed = urlparse(sharepoint_url)
        hostname = parsed.hostname
        if not hostname:
            raise ValueError("Invalid URL: no hostname found.")
        
        # Split the path and extract only the site path: /sites/<name> or /teams/<name>
        path_parts = parsed.path.strip("/").split("/")

        # Look for the first occurrence of "sites" or "teams"
        if "sites" in path_parts:
            idx = path_parts.index("sites")
        elif "teams" in path_parts:
            idx = path_parts.index("teams")
        else:
            raise ValueError("URL does not contain a /sites/ or /teams/ path.")

        # Rebuild only the site-level portion, ignore page/file paths
        site_path = "/".join(path_parts[idx:idx+2])

        # Construct Graph API format
        graph_url = f"{hostname}:/{site_path}"
        
        return graph_url
    
