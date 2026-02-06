import requests
import base64
import os

from dotenv import load_dotenv

# Initialize the .env file
load_dotenv()
# Environment Variables
API_KEY = os.getenv("API_KEY")
API_URL = os.getenv("API_URL")
LOCATION = os.getenv("LOCATION")
RESPONDER_ID = os.getenv("RESPONDER_ID")
REQUESTER_ID = os.getenv("REQUESTER_ID")


class FreshServiceAPI:
    def __init__(self, api_url, api_key):
        self.ticket_api_url = api_url
        self.api_key = api_key
        self.encoded_api_key = base64.b64encode(
            f"{api_key}:x".encode()).decode()
        self.headers = {
            "Authorization": f"Basic {self.encoded_api_key}",
            "Content-Type": "application/json",
        }

    def _build_url(self, endpoint):
        """Help to build a full URL"""
        return f"{API_URL}/{endpoint.lstrip('/')}"

    def update_ticket_resolution(self, ticket_id, resolution):
        url = self._build_url(f"tickets/{ticket_id}")
        response = requests.put(
            url,
            headers = self.headers,
            json = {"resolution_notes": resolution},
        )
        # print("API Response:", response.json())
        if response.status_code == 200:
            print(" ✅- Resolution notes updated successfully!")
            return response.json()
        else:
            print(
                f"⚠️ - Failed to update resolution notes: {response.status_code}")
            print("Response: ", response.json())
            return None

    def create_ticket(
            self,
            subject,
            description = None,
            email = None,
            category = None,
            priority = None,
            status = None,
            resolution_notes = None,
            type = None,
            requester_id = None,
            responder_id = None,
            group_id = None,
    ):
        url = self._build_url("tickets")
        response = requests.post(
            url,
            headers = self.headers,
            json = {
                "subject": subject,
                "description": description,
                "email": email,
                "priority": priority,  # 1: Low, 2: Medium, 3: High, 4: Urgent
                "status": status,  # 2: Open, 3: Pending, 4: Resolved, 5: Closed
                "resolution_notes": resolution_notes,
                "type": type,
                "requester_id": requester_id,
                "group_id": group_id,
                # Group where the ticket should be assigned
                "responder_id": responder_id,
                # Who the ticket should be assigned to
                "custom_fields": {
                    "please_select_the_service": category,
                    "site": os.getenv("SITE"),
                },
            },
        )

        # Check if the ticket response was good
        if response.status_code == 201:
            print(f"✅ - {type} ticket created successfully!")
            return response.json()
        elif response.status_code == 400:
            print("⚠️ - Failed to create ticket: Bad request (400)")
            print("Response: ", response.json())
            return None
        else:
            print(f"⚠️ - Failed to create ticket: {response.status_code}")
            return None


