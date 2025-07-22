
from dotenv import load_dotenv
load_dotenv()

import os
from flask import Flask, jsonify, request
import requests
from msal import ConfidentialClientApplication

import mimetypes
import fitz  # PyMuPDF
from docx import Document
from io import BytesIO


app = Flask(__name__)

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

def get_access_token():
    app_auth = ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    token_response = app_auth.acquire_token_for_client(scopes=SCOPE)
    return token_response

@app.route("/")
def home():
	return "Hello from your Copilot Flask backend!"

@app.route("/get-sites")
def get_sites():
	token_response = get_access_token()

	if "access_token" in token_response:
	    headers = {"Authorization": f"Bearer {token_response['access_token']}"}
	    graph_url = "https://graph.microsoft.com/v1.0/sites?search=*"
	    response = requests.get(graph_url, headers=headers)
	    return jsonify(response.json())
	else:
	    return jsonify({"error": "Could not acquire token", "details": token_response})

@app.route("/get-site-id")
def get_site_id():
    token_response = get_access_token()

    if "access_token" in token_response:
	    headers = {"Authorization": f"Bearer {token_response['access_token']}"}
	    site_url = "https://graph.microsoft.com/v1.0/sites/canalwin.sharepoint.com:/sites/TestSiteAI"
	    response = requests.get(site_url, headers=headers)
	    return jsonify(response.json())
    else:
	    return jsonify({"error": "Could not acquire token", "details": token_response})


def list_all_files(drive_id, folder_id=None, headers=None):
    files = []
    base_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
    endpoint = f"{base_url}/items/{folder_id}/children" if folder_id else f"{base_url}/root/children"

    response = requests.get(endpoint, headers=headers)
    if response.status_code == 200:
        items = response.json().get("value", [])
        for item in items:
            if item.get("folder"):
                # Recursively list files in subfolders
                files.extend(list_all_files(drive_id, item["id"], headers))
            else:
                files.append({
                    "name": item.get("name"),
                    "id": item.get("id"),
                    "webUrl": item.get("webUrl"),
                    "lastModifiedDateTime": item.get("lastModifiedDateTime"),
                    "size": item.get("size")
                })
    return files

@app.route("/list-files")
def list_files():
    token_response = get_access_token()

    if "access_token" in token_response:
        headers = {"Authorization": f"Bearer {token_response['access_token']}"}

        # Replace with your actual site ID
        site_id = "canalwin.sharepoint.com,40ae91cc-e81a-43d7-b21c-cff0110b85b8,486e30a0-a11f-46ff-9c15-98583f506e92"
        site_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"

        # Get the drive ID for the "Documents" library
        drive_response = requests.get(site_url, headers=headers)
        if drive_response.status_code == 200:
            drives = drive_response.json().get("value", [])
            documents_drive = next((d for d in drives if d.get("name") == "Documents"), None)

            if documents_drive:
                all_files = list_all_files(documents_drive["id"], headers=headers)
                return jsonify(all_files)
            else:
                return jsonify({"error": "Documents library not found"})
        else:
            return jsonify({"error": "Failed to retrieve drives", "details": drive_response.json()})
    else:
        return jsonify({"error": "Could not acquire token", "details": token_response})


@app.route("/get-file-content")
def get_file_content():
    file_id = request.args.get("file_id")
    if not file_id:
        return jsonify({"error": "Missing file_id parameter"}), 400

    token_response = get_access_token()
    if "access_token" not in token_response:
        return jsonify({"error": "Could not acquire token", "details": token_response}), 500

    headers = {"Authorization": f"Bearer {token_response['access_token']}"}

    
# Step 1: Get the drive ID dynamically (same logic as in /list-files)
    site_id = "canalwin.sharepoint.com,40ae91cc-e81a-43d7-b21c-cff0110b85b8,486e30a0-a11f-46ff-9c15-98583f506e92"
    site_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drive_response = requests.get(site_url, headers=headers)

    if drive_response.status_code != 200:
        return jsonify({"error": "Failed to retrieve drives", "details": drive_response.json()}), drive_response.status_code

    drives = drive_response.json().get("value", [])
    documents_drive = next((d for d in drives if d.get("name") == "Documents"), None)

    if not documents_drive:
        return jsonify({"error": "Documents library not found"}), 404

    drive_id = documents_drive["id"]

    # Step 2: Download the file
    download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
    file_response = requests.get(download_url, headers=headers)

    if file_response.status_code != 200:
        return jsonify({"error": "Failed to download file", "details": file_response.text}), file_response.status_code

    content_type = file_response.headers.get("Content-Type", "")
    file_bytes = BytesIO(file_response.content)

    # Step 3: Extract text based on file type
    try:
        if "pdf" in content_type:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text = "\n".join([page.get_text() for page in doc])
        elif "wordprocessingml" in content_type or file_id.endswith(".docx"):
            doc = Document(file_bytes)
            text = "\n".join([para.text for para in doc.paragraphs])
        else:
            return jsonify({
                "error": "Unsupported file type",
                "content_type": content_type
            }), 415
    except Exception as e:
        return jsonify({
            "error": "Failed to extract content",
            "details": str(e)
        }), 500

    return jsonify({
        "file_id": file_id,
        "content_type": content_type,
        "extracted_text": text
    })


if __name__ == "__main__":
	app.run(debug=True)
