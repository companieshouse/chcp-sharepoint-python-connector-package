# SharePoint Python Connection Package

This package provides a Python interface for interacting with Microsoft SharePoint via the Microsoft Graph API. It supports authentication, file upload/download and folder navigation.

- [SharePoint Python Connection Package](#sharepoint-python-connection-package)
  - [Features](#features)
  - [Installation](#installation)
  - [Instructions: Using the spconnect Package](#instructions-using-the-spconnect-package)
    - [1. Install Required Packages](#1-install-required-packages)
    - [2. Store Credentials in a .env File](#2-store-credentials-in-a-env-file)
    - [3. Authenticate with Azure AD Using Environment Variables](#3-authenticate-with-azure-ad-using-environment-variables)
    - [4. Get Site and Drive IDs](#4-get-site-and-drive-ids)
      - [4.1 site\_id](#41-site_id)
      - [4.2 drive\_id](#42-drive_id)
    - [5. List Folder Contents](#5-list-folder-contents)
    - [6. Download a File](#6-download-a-file)
    - [7. Load a file into memory (as a byte file)](#7-load-a-file-into-memory-as-a-byte-file)
    - [8. Upload a File](#8-upload-a-file)
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

You can install this package locally using [uv](https://github.com/astral-sh/uv):

To so this, simply run:

```sh
uv pip install </path/to/spconnect_package>
```
Replace `</path/to/spconnect_package>` with the path to this folder. uv will build and install the package automatically.


To install from the GitHub repo use:

```sh
uv pip install git+<github url>
```

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


### 4. Get Site and Drive IDs

The Microsoft Graph API uses a specific format to define the location of a file in a Sharepoint site. This can be summarised as:

`site_id/drive_id/file_id`

*Note: confusingly the `file_id` is a distinct identifier pointing to a file within a drive. a file can be within folders in the drive, but it will still have a distinct ID*

#### 4.1 site_id
The Microsoft Graph API uses a url in a specific format (not the url found in the browser when you visit a Sharepoint site)
`to_graph_site_url()` converts the url found in a browser to the graph api format. This allows you to visit a Sharepoint site in browser copy the url, and then paste it into the method to produce the graph api url. 

``` python
graph_site_url = client.to_graph_site_url("https://yourcompany.sharepoint.com/sites/YourSite")
site_id = client.get_site_id(graph_site_url)
```

#### 4.2 drive_id

To view the drives within a site, you can use: 

``` python
client.get_drives(site_id)
```

This will return a dictionary of drive names and IDs. 

If you know the name of a drive (e.g. you are looking at it in a browser), the drive id can also be found using the name of the drive.

```python
client.resolve_drive_id(site_id, "Documents")
```


### 5. List Folder Contents
To list the contents of a folder (root of a drive, or a subfolder within a drive):

```python
# for root of drive:
contents = client.get_folder_content(site_id, drive_id)
# For a subfolder:
contents = client.get_folder_content(site_id, drive_id, folder_path="Shared Documents/Reports")
```

### 6. Download a File
To download a file from SharePoint:

```python
client.download_file_to_disk(site_id, drive_id, file_id, local_path="./downloads")
```

you can get the file id the folder contents dictionary, or from it's name, and the name of the folder it's in (if it's in a folder).
you can use `resolve_file_id()` to get the file ID from it's name:

```python
file_id = client.resolve_file_id(site_id, drive_id, file_name="file.csv", folder_path="Shared Documents/Reports")
```

### 7. Load a file into memory (as a byte file)

Load a file into memory:

```python
byte_file = client.download_file_bytes(site_id, drive_id, file_id)

# Suppose the byte_file is an Excel file and you want to load it into polars
df = pl.read_excel(BytesIO(byte_file))
```

### 8. Upload a File
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
