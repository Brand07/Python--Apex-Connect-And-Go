# ruff: noqa: F405, F403
import polars as pl
from helium import *
from dotenv import load_dotenv
import os
import time
import sys
import logging
from exceptions import LoginError
import csv

# Configure logging to write to a file
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("apex.log"), logging.StreamHandler()],
)

# Load environment variables
load_dotenv()

# Initialize counters
users_added = 0
users_edited = 0


apex_users = pl.read_excel("New_Apex_Users.xlsx")
print(f"Number of users to be added: {len(apex_users)}\n")
print(f"Users to be added:\n {apex_users}\n")


def format_badge_number(badge_number):
    # Converts the badge number to a string
    badge_str = str(badge_number)
    # Check the length of the badge number
    if len(badge_str) == 4:
        # If 4 digits long, add a "0" at the beginning
        return "0" + badge_str
    elif len(badge_str) == 5:
        # If 5 digits long, return as is
        return badge_str
    elif len(badge_str) > 5:
        print("Badge number is too long. Please correct entries in the file.")
        logging.error("Badge number is too long. Please correct entries in the file.")
        sys.exit()
    else:
        print("Badge number must be 4 or 5 digits long")
        logging.error("Badge number must be 4 or 5 digits long")
        return None


def open_apex():
    # Start Firefox and navigate to the APEX URL
    start_firefox(os.getenv("APEX_URL"))  # noqa: F405

    # Wait for the page to load
    wait_until(Button("Sign In »").exists)  # noqa: F405
    highlight("Username")  # noqa: F405
    write(os.getenv("APEX_USERNAME"), into="Username")
    highlight("Password")
    write(os.getenv("APEX_PASSWORD"), into="password")
    #write("Christmas77", into="Password")
    click(Button("Sign In »"))  # noqa: F405
    if Text("Invalid Login/Password!").exists():
        raise LoginError("Invalid username or password. Please try again.")

    wait_until(Link("Profile Manager").exists)

    # Open the profile manager
    open_profile_manager()


def open_profile_manager():
    highlight(Link("Profile Manager"))
    click(Link("Profile Manager"))
    highlight(Link("Manage Users"))
    click(Link("Manage Users"))
    print("Proceeding to add users")
    time.sleep(1)
    process_users()


def process_users():
    global first_name, last_name, employee_id, badge_num, department
    # Loop through each row in the Excel file
    print("Adding users to the system.")
    for row in apex_users.iter_rows(named=True):
        first_name = row["First Name"]
        last_name = row["Last Name"]
        employee_id = row["Badge Number"]
        badge_num = int(row["Badge Number"])
        department = row["Department"]

        add_user(first_name, last_name, employee_id, badge_num, department)

     # Write the counters to a CSV file
    with open('user_tally.csv', mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['Users Added', 'Users Edited'])
        writer.writerow([users_added, users_edited])


def add_user(first_name, last_name, employee_id, badge_num, department):
    """Adds the user to the system. If the badge number is already in the system,
    it will edit the existing user info with the new information."""

    global users_added, users_edited

    print("Waiting for 'Add a User' link to exist.")
    wait_until(Link("Add a User").exists)
    highlight(TextField(below=Link("Add a User")))
    click(TextField(below=Link("Add a User")))
    write("", into=TextField(below=Link("Add a User")))
    write(f"{badge_num}")

    highlight(Button("Search"))
    click(Button("Search"))

    user_element = S("//*[@id='tr0']")
    time.sleep(1)

    if user_element.exists():
        print(f"{badge_num} already exists. Chanigng the existing info.")
        last_name_element = S("#tr0 > td:nth-child(1) > a:nth-child(1)")
        highlight(last_name_element)
        click(last_name_element)

        first_name_field = TextField(to_right_of=Text("First Name *:"))
        highlight(first_name_field)
        click(first_name_field)
        write("", into=first_name_field)  # Clears the field
        write(first_name, into=first_name_field)  # Clears the field

        last_name_field = TextField(to_right_of=Text("Last Name *:"))
        highlight(last_name_field)  # noqa: F405
        click(last_name_field)
        write("", into=last_name_field)  # Clears the field
        write(last_name, into=last_name_field)

        emp_id_field = TextField(to_right_of=Text("Employee ID *:"))
        highlight(emp_id_field)
        write("", into=emp_id_field)
        write(employee_id, into=emp_id_field)

        badge_number_field = TextField(to_right_of=Text("Badge #:"))
        highlight(badge_number_field)
        write("", into=badge_number_field)
        write(format_badge_number(badge_num))

        dept = Link("User Group Membership")
        click(dept)
        edit_all_checkboxes()
        # click(Button("Save"))
        edit_group_assignment(department)
        users_edited += 1
        logging.info(f"Badge # {badge_num} edited.")
        print("Clicking the save button")

    else:
        print("Badge number doesn't exist -- adding user.")
        logging.info(f"Badge # {badge_num} doesn't exist -- adding user.")
        add_user_link = Link("Add a User")
        click(add_user_link)

        first_name_field = TextField(to_right_of=Text("First Name *:"))
        click(first_name_field)
        write("", into=first_name_field)  # Clears the field
        write(first_name)  # Need to add logic to add the first name

        last_name_field = TextField(to_right_of=Text("Last Name *:"))
        click(last_name_field)
        write("", into=last_name_field)
        write(last_name)  # Need to add logic to add the last name

        emp_id_field = TextField(to_right_of=Text("Employee ID *:"))
        write("", into=emp_id_field)
        write(employee_id)  # placeholder

        badge_number_field = TextField(to_right_of=Text("Badge #:"))
        write("", into=badge_number_field)
        write(format_badge_number(badge_num))  # placeholder

        users_added += 1

        dept = Link("User Group Membership")
        click(dept)
        uncheck_all_checkboxes()
        group_assignment(department)


def edit_all_checkboxes():
    wait_until(Text("User Group Membership").exists)
    print("Unchecking all checkboxes.")
    checkboxes = [
        "input[id='editMembershipCheck0']",
        "input[id='editMembershipCheck1']",
        "input[id='editMembershipCheck2']",
        "input[id='editMembershipCheck3']",
        "input[id='editMembershipCheck4']",
        "input[id='editMembershipCheck5']",
        "input[id='editMembershipCheck6']",
        "input[id='editMembershipCheck7']",
        "input[id='editMembershipCheck8']",
        "input[id='editMembershipCheck9']",
        "input[id='editMembershipCheck10']",
        "input[id='editMembershipCheck11']",
    ]

    for checkbox in checkboxes:
        checkbox_element = S(checkbox).web_element
        if checkbox_element.is_selected():
            highlight(S(checkbox))
            click(S(checkbox))
        else:
            print("Checkbox is already unchecked.")

def uncheck_all_checkboxes():
    """Unchecks all the cheboxes when adding a new user"""
    wait_until(Text("User Group Membership").exists)
    print("Unchecking all checkboxes.")
    checkboxes = [
        "input[id='membershipCheck0']",
        "input[id='membershipCheck1']",
        "input[id='membershipCheck2']",
        "input[id='membershipCheck3']",
        "input[id='membershipCheck4']",
        "input[id='membershipCheck5']",
        "input[id='membershipCheck6']",
        "input[id='membershipCheck7']",
        "input[id='membershipCheck8']",
        "input[id='membershipCheck9']",
        "input[id='membershipCheck10']",
        "input[id='membershipCheck11']",
    ]

    for checkbox in checkboxes:
        checkbox_element = S(checkbox).web_element
        if checkbox_element.is_selected():
            highlight(S(checkbox))
            click(S(checkbox))
        else:
            print("Checkbox is already unchecked.")


def group_assignment(department):
    print("Assigning the group.")
    if department == "Cycle Count":
        return group_selection(2)
    elif department == "General":
        return group_selection(3)
    elif department == "Material Handler":
        return group_selection(4)
    elif department == "Sort":
        return group_selection(5)
    elif department == "Voice Pick":
        return group_selection(6)


def edit_group_assignment(group):
    """HTML IDs are different when editing a user
    vs adding a user."""
    print("Editing the group assignment.")
    time.sleep(1)
    if department == "Cycle Count":
        return edit_group_selection(2)
    elif department == "General":
        return edit_group_selection(3)
    elif department == "Material Handler":
        return edit_group_selection(4)
    elif department == "Sort":
        return edit_group_selection(5)
    elif department == "Voice Pick":
        return edit_group_selection(6)


def edit_group_selection(group):
    try:
        wait_until(Text("User Group Membership").exists)
        print("Edit window is open.")
        checkbox = f"input[id='editMembershipCheck{group}']"
        #checkbox_element = S(checkbox).web_element
        highlight(S(checkbox))
        click(S(checkbox))
        click(Button("Save"))
        click(Button("Save"))
    except Exception as e:
        print(e)


def group_selection(group):
    try:
        wait_until(Text("User Group Membership").exists)
        checkbox = f"input[id='membershipCheck{group}']"
        #checkbox_element = S(checkbox).web_element
        print("Checkbox is already selected.")
        highlight(S(checkbox))
        click(S(checkbox))
        click(Button("Add"))
        click(Button("Submit"))
        wait_until(Button("Ok").exists)
        click(Button("Ok"))
    except Exception as e:
        print(e)


open_apex()
print(f"{users_added} users added.")
print(f"{users_edited} users edited.")
