from helium import *
from dotenv import load_dotenv
import os

# Load environment variables
load_dotenv()



def open_apex():
    # Start Firefox and navigate to the APEX URL
    start_firefox(os.getenv('APEX_URL'))

# Wait for the page to load
    wait_until(Button("Sign In »").exists)
    write(os.getenv('APEX_USERNAME'), into="Username")
    write(os.getenv('APEX_PASSWORD'), into="Password")
    click(Button("Sign In »"))


open_apex()