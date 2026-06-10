from helium import *
from dotenv import load_dotenv
import os
import time
import sys
import logging
from exceptions import LoginError, GroupAssignmentException
import pandas as pd
from fresh_import import FreshServiceAPI, REQUESTER_ID, RESPONDER_ID

# Control the ticket logging by the flag
LogTickets = True

# Configure logging to write to a file
logging.basicConfig(
    level=loggin.info,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("apex.log"), logging.StreamHandler()],
)

# Load the environment variables
load_dotenv()