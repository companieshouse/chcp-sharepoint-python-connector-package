# SharePoint Python Connection Package

This package provides a Python interface for interacting with Microsoft SharePoint via the Microsoft Graph API. It supports authentication, file upload/download and folder navigation.

- [SharePoint Python Connection Package](#sharepoint-python-connection-package)
  - [Features](#features)
  - [Installation](#installation)
  - [Instructions: Using the spconnect Package](#instructions-using-the-spconnect-package)
    - [1. Install Required Packages](#1-install-required-packages)
    - [2. Store Credentials in a .env File](#2-store-credentials-in-a-env-file)
    - [3. Authenticate with Azure AD Using Environment Variables](#3-authenticate-with-azure-ad-using-environment-variables)
    - [4 Get Site, Drive and File IDs](#4-get-site-drive-and-file-ids)
      - [4.1 Extract IDs from a SharePoint File URL](#41-extract-ids-from-a-sharepoint-file-url)
      - [4.2 Find the `site_id`](#42-find-the-site_id)
      - [4.3 Find the `drive_id`](#43-find-the-drive_id)
      - [4.4 Find the `file_id`](#44-find-the-file_id)
    - [5. Download a File](#5-download-a-file)
    - [6. Load a file into memory (as a byte file)](#6-load-a-file-into-memory-as-a-byte-file)
    - [7. Upload a File](#7-upload-a-file)
  - [Notes](#notes)
  - [Microsoft Graph API resources](#microsoft-graph-api-resources)
  - [License](#license)


## Features
- Authenticate with Microsoft Entra ID (tenant ID, client ID, client secret)
- Retrieve site and drive IDs
- List folder contents
- Download files to disk or as bytes
- Upload files from disk
- Convert SharePoint URLs to Graph API site identifiers

## Installation

You can install this package directly from the GitHub repo using [uv](https://github.com/astral-sh/uv). 

In the terminal type:

```sh
uv add git+https://github.com/companieshouse/chcp-sharepoint-python-connector-package.git
```

You can also install it from a local clone of the repository using:

```sh
uv add </path/to/spconnect_package>
```
Replace `</path/to/spconnect_package>` with the path to this folder. uv will build and install the package automatically.



## Instructions: Using the spconnect Package

This guide provides step-by-step instructions for using the `spconnect` package to interact with SharePoint via the Microsoft Graph API. The steps below use generic SharePoint site names, drives, folders, and files so you can adapt them to your own Sharepoint Site.

### 1. Install Required Packages
Ensure you have installed the `spconnect` package.

### 2. Store Credentials in a .env File
Create a `.env` file in the root directory containing your Microsoft Entra ID credentials:

```
TENANT_ID=<your-tenant-id>
CLIENT_ID=<your-client-id>
CLIENT_SECRET=<your-client-secret>
```

### 3. Authenticate with Azure AD Using Environment Variables
Load credentials from the environment using the `python-dotenv` package:

``` python
from dotenv import load_dotenv
import os
from spconnect import SharePointClient

load_dotenv()

tenant_id = os.getenv("TENANT_ID")
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")

client = SharePointClient(
    tenant_id=tenant_id,
    client_id=client_id,
    client_secret=client_secret,
    scopes=["https://graph.microsoft.com/.default"]
)
```

*Note: Make sure to install `python-dotenv` if you haven't already:*

### 4 Get Site, Drive and File IDs

The Microsoft Graph API uses a specific format to define the location of a file in a Sharepoint site. This can be summarised as:

`site_id/drive_id/file_id`

*Note: confusingly the `file_id` is a distinct identifier pointing to a file within a drive. a file can be within folders in the drive, and the file ID covers the folders and file name*

#### 4.1 Extract IDs from a SharePoint File URL

You can use the `parse_url_to_ids()` method to extract the `site_id`, `drive_id`, and `file_id` directly from a SharePoint file URL. This is useful if you have a link to a file and want to quickly get the identifiers needed for other API calls. 

To get the file URL from SharePoint:
- Locate the file in SharePoint.
- Click on the `More Actions` button (three dots: `...`) next to the file name.
- Scroll down and select `Path`.
- Click on `Copy Path` to copy the file URL to your clipboard.

```python
# Example SharePoint file URL (replace with your actual file URL)
file_url = "https://yourtenant.sharepoint.com/sites/YourSite/YourDrive/Your%20File.csv"

# This will return a dictionary with keys: 'site_id', 'drive_id', 'file_id'
ids = client.parse_url_to_ids(file_url)
print(ids)
# Output: {'site_id': '...', 'drive_id': '...', 'file_id': '...'}

# You can then use these IDs in other methods:
client.download_file_to_disk(ids['site_id'], ids['drive_id'], ids['file_id'], local_path="./downloads")
```

#### 4.2 Find the `site_id`
The Microsoft Graph API uses a url in a specific format (not the url found in the browser when you visit a Sharepoint site)
`to_graph_site_url()` converts the url found in a browser to the graph api format. This allows you to visit a Sharepoint site in browser, copy the url, and then paste it into the method to produce the graph api url. 

``` python
graph_site_url = client.to_graph_site_url("https://yourcompany.sharepoint.com/sites/YourSite")
site_id = client.get_site_id(graph_site_url)
```

#### 4.3 Find the `drive_id`

To view the drives within a site, you can use: 

``` python
client.get_drives(site_id)
```

This will return a dictionary of drive names and IDs. 

If you know the name of a drive (e.g. you are looking at it in a browser), the drive id can also be found using the name of the drive.

```python
client.resolve_drive_id(site_id, "Documents")
```

#### 4.4 Find the `file_id`

There are two main ways to find the `file_id` for a file in SharePoint:

**From folder contents:**
  Use `get_folder_content(site_id, drive_id, folder_path)` to list all files and folders in a location. This returns a dictionary where the keys are file IDs and the values are file names.

  ```python
  contents = client.get_folder_content(site_id, drive_id, folder_path="Shared Documents/Reports")
  ```

**From file name:**
  Use `resolve_file_id(site_id, drive_id, file_name, folder_path)` to get the file ID directly if you know the file name and (optionally) the folder path.

  ```python
  file_id = client.resolve_file_id(site_id, drive_id, file_name="file.csv", folder_path="Shared Documents/Reports")
  ```

You can then use the `file_id` with download, upload, or other file operations.

### 5. Download a File
To download a file from SharePoint:

```python
client.download_file_to_disk(site_id, drive_id, file_id, local_path="./downloads")
```

### 6. Load a file into memory (as a byte file)

Load a file into memory, this is useful if you want to perform operations on the data
before saving e.g. data cleaning:

```python
byte_file = client.download_file_bytes(site_id, drive_id, file_id)

# Suppose the byte_file is an Excel file and you want to load it into polars
df = pl.read_excel(BytesIO(byte_file))
```

### 7. Upload a File
To upload a file to SharePoint:
- **local_file_path**: Path to the file on your computer
- **upload_folder_path**: Target folder in SharePoint (optional: if not supplied, the file will be uploaded to the root of the drive)

```python
client.upload_file(
    local_file_path="./data/report.xlsx",
    site_id=site_id,
    drive_id=drive_id,
    upload_folder_path="Shared Documents/Reports"
)
```

## Notes
- Replace all example values with your actual SharePoint site, drive, folder, and file names.
- Ensure your Azure AD app has the necessary permissions for Microsoft Graph API.
- For large files (>3MB), the package automatically uses chunked upload.

---

## Microsoft Graph API resources

Microsoft have some great resources explaining how to use their Graph API

[Microsoft Learn](https://learn.microsoft.com/en-us/graph/use-the-api): High level explainer.

[Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer):
This site contains a cheat sheet of different API calls, and lets you test them.


## License
MIT License
