from helium import *
from dotenv import load_dotenv
import os
import time
import sys
import logging
import yaml
from exceptions import LoginError
import pandas as pd
from fresh_import import FreshServiceAPI, REQUESTER_ID, RESPONDER_ID

# Control the ticket logging by the flag
LogTickets = True

# Configure logging to write to a file
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("apex.log"), logging.StreamHandler()],
)

# Load the environment variables
load_dotenv()

# ==== Environment Variables =====

API_KEY = os.getenv("API_KEY")
API_URL = os.getenv("API_URL")
REQUESTER_ID = os.getenv("REQUESTER_ID")
RESPONDER_ID = os.getenv("RESPONDER_ID")
GROUP_ID = os.getenv("GROUP_ID")
fresh_service_api = FreshServiceAPI(api_url = API_URL, api_key = API_KEY)
APEX_USERS_URL = os.getenv("APEX_USERS_URL")
APEX_URL = os.getenv("APEX_URL")
APEX_USERNAME = os.getenv("APEX_USERNAME")
APEX_PASSWORD = os.getenv("APEX_PASSWORD")
excel_file = os.getenv("EXCEL_FILE")

# ==== END Environment Variables ====

# ==== Web Element Realted ====

def load_elements(path="elements.yaml"):
    with open(path, "r") as f:
        return yaml.safe_load(f)

web_elements = load_elements()

# ==== END Web Element Code ====

#Initialize Counters
users_added = 0
users_edited = 0

# ==== Excel Realted Info =====

apex_users = pd.read_excel(excel_file, sheet_name="Sheet1")
print(f"Users to be added: {len(apex_users)}\n")
print(f"Users to be added: \n {apex_users}\n")

# Check if the file is empty
if apex_users.empty:
    print("The excel file is empty. Make sure the file has user data in it.")
    logging.error("The Excel file is empty.")
    sys.exit()

# Archive sheet for users processed
archive_df = pd.read_excel(excel_file, sheet_name="Archive")

# ==== END Excel Realted Info ====

def format_badge_number(badge_number):
    # Convert the badge number to a string
    badge_str = str(badge_number)
    
    # Check the legnth of the badge number
    if len(badge_str) == 4:
        print("Appending a '0' to the badge number.")
        return "0" + badge_str
    # If the badge is already 5 digits long, return as is
    elif len(badge_str) == 5:
        print("No badge formatting needed.")
        return badge_str

    elif len(badge_str) > 5:
        print("Badge number is too long! Please double check entries in the file.")
        #TODO add logic to skip over the user instead of stopping the whole script
        sys.exit()
    
    else:
        print("Badge number must be 4 or 5 digits long.")
        logging.error("Badge number must be 4 or 5 digits long.")
        return None

def open_apex():
    """ Opens the Apex Website """
    try:
        start_firefox(APEX_URL)
    except Exception as e:
        return e

def login():
    # Controls the login logic
    try:
        # Wait for the page to load
        wait_until(S(web_elements["login_page"]["sign_in_button"]).exists)
        print("Sign In Button Exists.")
        # Send the username
        write(APEX_USERNAME, into=S(web_elements["login_page"]["username_field"]))

        # Send the password
        write(APEX_PASSWORD, into=S(web_elements["login_page"]["password_field"]))

    except Exception as e:
        print(e)


open_apex()