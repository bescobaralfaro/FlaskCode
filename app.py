
from dotenv import load_dotenv
load_dotenv()

import os
from flask import Flask, jsonify
from msal import ConfidentialClientApplication
import requests

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

@app.route("/list-files")
def list_files():
    token_response = get_access_token()

    if "access_token" in token_response:
	    headers = {"Authorization": f"Bearer {token_response['access_token']}"}
	    site_id = "canalwin.sharepoint.com,40ae91cc-e81a-43d7-b21c-cff0110b85b8,486e30a0-a11f-46ff-9c15-98583f506e92"
	    files_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root/children"
	    response = requests.get(files_url, headers=headers)
	    return jsonify(response.json())
    else:
	    return jsonify({"error": "Could not acquire token", "details": token_response})

if __name__ == "__main__":
	port = int(os.environ.get("PORT", 10000))
	app.run(host = "0.0.0.0", port = port)
