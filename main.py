from helium import *
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
    click(Link('Profile Manager'))



    

open_apex()