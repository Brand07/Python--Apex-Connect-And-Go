from helium import *
from helium import CLEAR
from dotenv import load_dotenv
import os
import pandas as pd

# Load environment variables
load_dotenv()

apex_users = os.getenv("EXCEL_FILE")

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
    write('Test')
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
        write("Working")
        

open_apex()