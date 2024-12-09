import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import getpass
import time
import sys



# Read data from sheet 1
apex_users_sheet_one = pd.read_excel("New_Apex_Users.xlsx", sheet_name="Sheet1")
# Define the excel sheet for new users
apex_users = pd.read_excel("New_Apex_Users.xlsx")

# Define the global webdriver
driver = webdriver.Firefox()
# Set the window size to 1920x1080
driver.set_window_size(1920, 1080)

def format_badge_number(badge_number):
    # Convert the badge number to a string
    badge_str = str(badge_number)
    
    # Check the length of the badge number
    if len(badge_str) == 4:
        # If 4 digits long, add a "0" at the beginning
        return "0" + badge_str
    elif len(badge_str) == 5:
        # If 5 digits long, return it as is
        return badge_str
    else:
        # Optionally, handle badge numbers that are not 4 or 5 digits long
        print("Badge number must be 4 or 5 digits long.")
        return None



def login_to_apex(retry_count=0):
    """
    Opens the webdriver, prompts for login info,
    and navigates to the User Management page.
    """
    try:
        #define the URL to connect load
        driver.get("https://apexconnectandgo.com/")
        #find the username field
        userID_element = driver.find_element(By.ID, "user.login_id")
        #clear the field in case it contains something
        userID_element.clear()
        #pass the user ID
        userID_element.send_keys(input("Enter your username: "))
        userID_element.send_keys(Keys.TAB)
        #locate the password field element
        pw_element = driver.find_element(By.ID, "user.password")
        #clear the field
        pw_element.clear()
            #call for the password input
        pw_element.send_keys(getpass.getpass("Enter your password: "))
        pw_element.send_keys(Keys.RETURN)
        #Give the page time to load
        time.sleep(4)
        # This looks for the INBOX element on the screen after a successful login
        if not driver.find_elements(By.XPATH, "/html/body/div[1]/div[3]/div/div[1]/div/h2"):
            raise Exception("Login failed. Dashboard element not found.")
        
        print("Navigating to the 'Manage Users' screen.")
        driver.get("https://apexconnectandgo.com/APEX-Login/accountAction_initManageuser.action?isShow=users")
    except Exception as e:
        print(f"An error occurred: {e}")
        if retry_count < 3:  # Set retry limit to 3 attempts
            print("Retrying login...")
            login_to_apex(retry_count + 1)
        else:
            print("Failed to login after multiple attempts. Exiting.")
            sys.exit(1)  # Exit the script with an error code

        print("Navigating to the 'Manage Users' screen.")
        #Navigate to the 'Manage Users' screen
        driver.get("https://apexconnectandgo.com/APEX-Login/accountAction_initManageuser.action?isShow=users")
        time.sleep(2)
    process_users()


    
    

def process_users():
    global first_name, last_name, employ_id, badge_num, department
    # Loop through each row in Excel and add users
    for index, row in apex_users.iterrows():
        first_name = row["First Name"]
        last_name = row["Last Name"]
        employ_id = row["Badge Number"]
        
        # Check if 'Badge Number' is NaN and handle accordingly
        if pd.isna(row["Badge Number"]):
            # Option 1: Skip this user
            continue
            # Option 2: Use a default value, e.g., 0 or some placeholder
            # badge_num = 0
            # Option 3: Handle NaN differently as per your requirements
        else:
            badge_num = int(row["Badge Number"])
        
        department = row["Department"]
        
        add_user(first_name, last_name, employ_id, badge_num, department)


def add_user(first_name, last_name, employee_id, badge_num, department):
    time.sleep(2)
    print(f"Searching For Badge Number  {badge_num}.")
    existing_user_search = driver.find_element(By.ID, "searchUsersText")
    existing_user_search.click()
    existing_user_search.clear()
    existing_user_search.send_keys(badge_num)
    existing_user_search.send_keys(Keys.RETURN)
    #check for element on the page
    time.sleep(2)
    user_element = driver.find_elements(By.XPATH, "//*[@id='tr0']")
    #existing_user_search.clear()
    time.sleep(1)
    #check if the element exists
    if len(user_element) == 1:

        print(f"{badge_num} already exists. Proceeding to change existing info.")
        last_name_element = driver.find_element(By.CSS_SELECTOR, '#tr0 > td:nth-child(1) > a:nth-child(1)')
        last_name_element.click()
        time.sleep(2)
        f_name = driver.find_element(By.ID, "edit_user.first_name")\
        # Clear the 'First Name' Field
        f_name.clear()
        f_name.send_keys(first_name)
        l_name = driver.find_element(By.ID, "edit_user.last_name")
        # Clear the 'Last Name' Field
        l_name.clear()
        l_name.send_keys(last_name)
        emp_id = driver.find_element(By.ID, "employeeId")
        # Clear the 'Employee ID' field.
        emp_id.clear()
        emp_id.send_keys(employee_id)
        badge_number = driver.find_element(By.ID, "badgeNumber")
        # Clear the 'Badge Number' field
        badge_number.clear()
        #badge_number.send_keys("0", badge_num)
        badge_number.send_keys(format_badge_number(badge_num))
        dept = driver.find_element(By.LINK_TEXT, "User Group Membership:")
        dept.click()
        time.sleep(1)
        edit_all_checkboxes()
        time.sleep(1)
        edit_group_assignment(department)
        time.sleep(1)
        print('Clicking the save button')
        
        try:
            save_button = driver.find_element(By.XPATH, '/html/body/div[21]/div[3]/div/button[2]')
            ActionChains(driver).move_to_element(save_button).click().perform()
        except Exception:
            save_button = driver.find_element(By.XPATH, '/html/body/div[21]/div[11]/div/button[2]')
            ActionChains(driver).move_to_element(save_button).click().perform()
        
        print("Save button clicked.")
        time.sleep(2)
        submit = driver.find_element(By.ID, 'updateUser')
        submit.click()
        print("Changes saved.")
        print(f"User {first_name} {last_name} has been added with {department} permissions.")
    else: 
        print(f"{badge_num} doesn't exist. Adding now..")
        add_user_link = driver.find_element(By.ID, "addUserLink")
        add_user_link.click()
        f_name = driver.find_element(By.ID, "user.first_name")
        f_name.send_keys(first_name)
        l_name = driver.find_element(By.ID, "user.last_name")
        l_name.send_keys(last_name)
        emp_id = driver.find_element(By.ID, "addPassport.employee_id")
        emp_id.send_keys(employee_id)
        badge_number = driver.find_element(By.ID, "addPassport.user_card_key")
        #badge_number.send_keys("0", badge_num)
        badge_number.send_keys(format_badge_number(badge_num))
        dept = driver.find_element(By.LINK_TEXT, "User Group Membership:")
        dept.click()
        time.sleep(1)
        uncheck_all_checkboxes()
        time.sleep(1)
        group_assignment(department)
        add_button = driver.find_element(By.XPATH, "//button[normalize-space()='Add']")
        ActionChains(driver).move_to_element(add_button).click().perform()
        time.sleep(1)
        submit = driver.find_element(By.XPATH, "//button[normalize-space()='Submit']")
        submit.click()
        popup = driver.find_element(By.XPATH, "/html/body/div[4]/div[3]/div/button")
        popup.click()
        print(f"User {first_name} {last_name} has been added with {department} permissions.")
        
        


def edit_all_checkboxes():
    """
    Specifically unchecks all of the checkboxes when you go to edit a user.
    Editing a user entails different XPATH IDs than adding a user.
    This function loops through all of the checkboxes and unchecks them.
    
    """
    time.sleep(1)
    print("Unchecking all checkboxes.")
    checkboxes = ['//*[@id="editMembershipCheck0"]', '//*[@id="editMembershipCheck1"]', '//*[@id="editMembershipCheck2"]',
                    '//*[@id="editMembershipCheck3"]', '//*[@id="editMembershipCheck4"]', '//*[@id="editMembershipCheck5"]',
                    '//*[@id="editMembershipCheck6"]', '//*[@id="editMembershipCheck7"]', '//*[@id="editMembershipCheck8"]',
                    '//*[@id="editMembershipCheck9"]', '//*[@id="editMembershipCheck10"]', '//*[@id="editMembershipCheck11"]',]
    

    for xpath in checkboxes:
        try:
            # Wait for the checkbox to be clickable
            checkbox_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath)))
            # Optional: Check if a specific state is desired (e.g., checked/unchecked)
            if checkbox_element.is_selected():
                checkbox_element.click()
        except TimeoutException:
            print(f"Checkbox with XPath '{xpath}' not found or not clickable.")
            
def uncheck_all_checkboxes():
    time.sleep(1)
    print("Unchecking all checkboxes.")
    checkboxes = ['//*[@id="membershipCheck0"]', '//*[@id="membershipCheck1"]', '//*[@id="membershipCheck2"]',
                  '//*[@id="membershipCheck3"]', '//*[@id="membershipCheck4"]', '//*[@id="membershipCheck5"]',
                  '//*[@id="membershipCheck6"]', '//*[@id="membershipCheck7"]', '//*[@id="membershipCheck8"]',
                  '//*[@id="membershipCheck9"]', '//*[@id="membershipCheck10"]', '//*[@id="membershipCheck11"]',]
    
    for xpath in checkboxes:
        try:
            # Wait for the checkbox to be clickable
            checkbox_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath)))
        
            # Optional: Check if a specific state is desired (e.g., checked/unchecked)
            if checkbox_element.is_selected():
                checkbox_element.click()
        except TimeoutException:
            print(f"Checkbox with XPath '{xpath}' not found or not clickable.")


def group_assignment(group):
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
    """
    HTML IDs are different when editing a user
    VS when adding a user.
    """
    
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
        # Wait for the checkbox to be clickable
        checkbox = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, f"editMembershipCheck{group}")))
    
        # Optional: Check if a specific state is desired (e.g., checked/unchecked)
        if not checkbox.is_selected():
            checkbox.click()
    except TimeoutException:
        print(f"Checkbox with ID 'editMembershipCheck{group}' not found or not clickable.")
    
def group_selection(group):
    emp_group = driver.find_element(By.LINK_TEXT, "User Group Membership:")
    emp_group.click()
    time.sleep(1)
    xpath = f"//input[@id='membershipCheck{group}']"
    checkbox = driver.find_element(By.XPATH, xpath)
    checkbox.click()


        
    





login_to_apex()


