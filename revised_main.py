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
SDR_PERM = os.getenv("SDR_PERM")

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
    badge_str = str(badge_number)
    
    if len(badge_str) == 4:
        print("Appending a '0' to the badge number.")
        return "0" + badge_str
    elif len(badge_str) == 5:
        print("No badge formatting needed.")
        return badge_str
    else:
        raise ValueError(f"Unexpected badge number length {len(badge_str)}: '{badge_str}'")

def open_apex():
    """ Opens the Apex Website """
    try:
        start_firefox(APEX_URL)
        # Wait for the page to load
        wait_until(S(web_elements["login_page"]["sign_in_button"]).exists)
        print("Sign In Button Exists.")

        login()

    except Exception as e:
        return e

def login():
    # Controls the login logic
    try:
        # Send the username
        write(APEX_USERNAME, into=S(web_elements["login_page"]["username_field"]))

        # Send the password
        write(APEX_PASSWORD, into=S(web_elements["login_page"]["password_field"]))

        # CLick the Sign in Button
        click(S(web_elements["login_page"]["sign_in_button"]))

        # Wait for the profile button to load
        wait_until(S(web_elements["to_user_profiles"]["profile_manager"]).exists)

    
        open_profile_manager()

    except Exception as e:
        print(e)

def open_profile_manager():
    print("Opening the profile manager")
    # Navigates to the profile page to search users
    try:
        print("Clicking on 'Profile Manager'")
        click(Link("Profile Manager"))

        print("Clicking on 'Manage Users'")
        click(Link("Manage Users"))

        search_users()
        

    except Exception as e:
        print(e)


def add_remove_sdr_permissons(first_name, last_name, employee_id, badge_num, department):
    """ Removes (NOT EDIT) SDR permissions if the user doesn't need them """
    try:
        if S(web_elements["other_elements"]["existing_sdr_perm"]) == SDR_PERM:
            print("User has SDR permissions.")
            #TODO add more logic to remove the permissions if necessary
        else:
            print("User does not have SDR permissions.")

    except Exception as e:
        return e

def add_add_sdr_permission(first_name, last_name, employee_id, badge_num, department):
    """Adds (Not EDIT) SDR permissions if the user doesn't need them"""
    try:
        print("Clicking on 'Rule Assignment'.")
        click(S(web_elements["add_user_page"]["add_user_rule_assignment"]))
        time.sleep(0.5)

        print("Clicking on the SDR permission.")
        click(S(web_elements["add_user_page"]["add_user_rule_assignment_sdr_perm"]))
        print("Saving the permission.")
        click(S(web_elements["add_user_page"]["add_user_rule_assignment_sdr_perm_save"]))

    except Except as e:
        return e

def search_users():
    global archive_df
    global apex_users
    global first_name, last_name, employee_id, badge_num, department

    print("Adding users to the system.")

    rows_to_move =[]

    for index, row in apex_users.iterrows():
        first_name = row["First Name"]
        last_name = row["Last Name"]
        employee_id = row["Badge Number"]
        badge_num = int(row["Badge Number"])
        department = row["Department"]

        try:
            print("Writing the badge number into the search box.")
            write(badge_num, into=S(web_elements["user_search"]["search_box"]))
            press(ENTER)

            # See if the badge number is already in the system
            if S(web_elements["user_search"]["existing_user"]).exists():
                print(f"{badge_num} - already exists... editing the user.")
                #TODO add an edit user function

            else:
                #TODO add an add user function
                print(f"Adding user: {first_name} {last_name} - {badge_num}")
                add_user(first_name, last_name, employee_id, badge_num, department)

        except Exception as e:
            print(e)
        


def add_user(first_name, last_name, employee_id, badge_num, department):
    try:
        print("Clicking the 'Add a User' link.")
        click(S(web_elements["user_search"]["add_user_button"]))

        # Enter the first name
        write(first_name, into=S(web_elements["add_user_page"]["add_first_name_element"]))
        # Enter the last name
        write(last_name, into=S(web_elements["add_user_page"]["add_last_name_element"]))
        # Enter the formatted badge number into the employee ID and Badge Number field
        write(format_badge_number(badge_num), into=S(web_elements["add_user_page"]["add_user_emp_id"]))
        write(format_badge_number(badge_num), into=S(web_elements["add_user_page"]["add_user_badge_num"]))

        
    
    except Exception as e:
        print(e)

        """

        Need to determine whether the user has or needs SDR access here.
        Needing it means the user needs to be saved, then edited to add or remove.

        If the user needs SDR access, it can be done here in the same place.

        """
    try:
        print("Determining the user's needs SDR permissions...")

        if department == "SDR":
            # Add the permission via the 'Rule Assignment' option
            click(S(web_elements["add_user_page"]["add_user_rule_assignment"]))
            # Click the SDR checkbox
            click(S(web_elements["add_user_page"]["add_user_rule_assignment_sdr_perm"]))
            # Save the option
            click(S(web_elements["add_user_page"]["add_user_rule_assignment_sdr_perm_save"]))
            #TODO save the user profile
            #TODO create and resolve ticket at this point

        else:
            click(S(web_elements)["add_user_page"]["add_user_group_membership"])
            #TODO call function to determine the permission needed.
            


    except Exception as e:
        return e



open_apex()