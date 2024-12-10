# (Work in Progress)

# Apex User Management Automation

This project automates the process of adding and editing users in the Apex system using Python and the Helium library for browser automation.

## Prerequisites

- Python 3.6 or higher
- Firefox browser
- `pip` for managing Python packages

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/Brand07/Python--Apex-Connect-And-Go.git
    ```

2. Install the required Python packages:
    ```sh
    pip install -r requirements.txt
    ```

3. Create a `.env` file in the root directory with the following content:
    ```dotenv
    APEX_URL=www.apexconnectandgo.com
    APEX_USERNAME=your_username
    APEX_PASSWORD=your_password
    ```

4. Place the `New_Apex_Users.xlsx` file in the root directory. This file should contain the user data to be added or edited.

## Usage

Run the script to start the automation process:
```sh
python main.py
```

## Logging

Logs are written to both the console and a file named `apex.log` in the root directory. The log file contains information about the users added or edited and any errors encountered during the process.

## Exception Handling

Custom exceptions are defined in the `exceptions.py` file. The `LoginError` exception is raised if there is an issue with logging into the Apex system.

## Project Structure

- `main.py`: The main script that performs the automation.
- `exceptions.py`: Contains custom exception classes.
- `requirements.txt`: Lists the Python dependencies for the project.
- `.env`: Contains environment variables for the Apex URL, username, and password.
- `apex.log`: Log file for the automation process.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.
