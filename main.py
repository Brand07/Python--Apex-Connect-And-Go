from helium import *
from helium import CLEAR
from dotenv import load_dotenv
import os
import pandas as pd

# Load environment variables
load_dotenv()

apex_users = os.getenv("EXCEL_FILE")

def format_badge_number(badge_number):
    # Converts the badge number to a string
    badge_str = str(badge_number)
    # Check the length of the badge number
    if len(badge_number) == 4:
        # If 4 digits long, add a "0" at the beginning
        return "0" + badge_str
    elif len(badge_str) == 5:
        # If 5 digits long, return as is
        return badge_str
    else:
        print("Badge number must be 4 or 5 digits long")
        return None

def open_apex():
    # Start Firefox and navigate to the APEX URL
    start_firefox(os.getenv('APEX_URL'))

# Wait for the page to load
    wait_until(Button("Sign In »").exists)
    write(os.getenv('APEX_USERNAME'), into="Username")
    write(os.getenv('APEX_PASSWORD'), into="Password")
    click(Button("Sign In »"))
    wait_until(Link('Profile Manager').exists)

    # Open the profile manager
    open_profile_manager()

def open_profile_manager():
    highlight(Link('Profile Manager'))
    click(Link('Profile Manager'))
    highlight(Link('Manage Users'))
    click(Link('Manage Users'))
    add_user()


def add_users_to_system():
    global first_name, last_name, employee_id, badge_num, department
    # Loop through each row in the Excel file
    for index, row in apex_users.iterrows():
        first_name = row['First Name']
        last_name = row['Last Name']
        employee_id = row['Badge Number']

        if pd.isna(row['Badge Number']):
            continue
        else:
            badge_num = int(row['Badge Number'])

        department = row['Department']

        add_user(first_name, last_name, employee_id, badge_num, department)
    

def add_user():#first_name, last_name, employee_id, badge_number, department):
    highlight(TextField(below=Link('Add a User')))
    click(TextField(below=Link('Add a User')))
    write('Barrack Obama')
    highlight(Button('Search'))
    click(Button('Search'))

    user_element = S("//*[@id='tr0']")

    if user_element.exists():
        print("Info already exists")
        #print(f"{badge_num} already exists. Chanigng the existing info.")
        last_name_element = S("#tr0 > td:nth-child(1) > a:nth-child(1)")
        click(last_name_element)
        print("Last name element clicked.")

        first_name_field = TextField(to_right_of=Text('First Name *:'))
        click(first_name_field)
        write("", into=first_name_field) # Clears the field
        write("Working") # Need to add logic to add the first name

        last_name_field = TextField(to_right_of=Text('Last Name *:'))
        click(last_name_field)
        write("", into=last_name_field)
        write('Last Name') # Need to add logic to add the last name

        emp_id_field = TextField(to_right_of=Text('Employee ID *:'))
        write("", into=emp_id_field)
        write('Employee ID') # placeholder

        badge_number_field = TextField(to_right_of=Text('Badge #:'))
        write("", into=badge_number_field)
        write('Badge Number') # placeholder

        # TODO
        """Add Logic to edit the group membership"""

        dept = Link("User Group Membership")
        click(dept)
    else:
        print("Badge number doesn't exist -- adding user.")
        add_user_link = Link("Add a User")
        click(add_user_link)

        first_name_field = TextField(to_right_of=Text('First Name *:'))
        click(first_name_field)
        write("", into=first_name_field) # Clears the field
        write("Working") # Need to add logic to add the first name

        last_name_field = TextField(to_right_of=Text('Last Name *:'))
        click(last_name_field)
        write("", into=last_name_field)
        write('Last Name') # Need to add logic to add the last name

        emp_id_field = TextField(to_right_of=Text('Employee ID *:'))
        write("", into=emp_id_field)
        write('Employee ID') # placeholder

        badge_number_field = TextField(to_right_of=Text('Badge #:'))
        write("", into=badge_number_field)
        write('Badge Number') # placeholder

        dept = Link("User Group Membership")
        click(dept)
        
        #TODO
        # Implement function to edit the group assignment

    


        

open_apex()