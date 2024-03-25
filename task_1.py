"""Module documentation for user_credentials_scraper."""
# pylint: disable=W0621

from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

def scrape_user_credentials(url):
    """
    Scrape user credentials from the given URL.

    This function takes a URL as an input and returns a tuple containing two lists:
    `user_ids` and `password`. The `user_ids` list contains the user IDs of the users,
    and the `password` list contains the corresponding passwords.

    Args:
    - url (str): The URL from which user credentials will be scraped.

    Returns:
    - user_ids (list): A list of user IDs.
    - password (str): A string containing the password.
    """
    driver = webdriver.Chrome()
    driver.get(url)
    user_ids_element = driver.find_element(By.ID, "login_credentials")
    user_ids = user_ids_element.text.split('\n')[1:]

    password_element = driver.find_element(By.CSS_SELECTOR, ".login_password")
    password = password_element.text.split(':')[-1].strip()

    driver.quit()

    return user_ids, password

def write_to_excel(user_ids, password, output_file):
    """
    Write user credentials to an Excel file.

    This function takes a list of user IDs, a password, and the desired output file name,
    and writes the data into an Excel file.

    Args:
    - user_ids (list): A list of user IDs.
    - password (str): A string containing the password.
    - output_file (str): The desired output file name for the Excel file.

    No return value.
    """
    wb = openpyxl.Workbook()
    sheet_credentials = wb.active
    sheet_credentials.title = "User credentials"

    sheet_credentials['A1'] = "User ID"
    sheet_credentials['B1'] = "User Name"
    sheet_credentials['C1'] = "Password"

    for idx, user_id in enumerate(user_ids, start=2):
        sheet_credentials[f'A{idx}'] = idx - 1  # User ID
        sheet_credentials[f'B{idx}'] = user_id  # User Name
        sheet_credentials[f'C{idx}'] = password  # Password

    wb.save(output_file)

if __name__ == "__main__":
    URL = "https://www.saucedemo.com/v1/"
    OUTPUT_FILE = "user_credentials.xlsx"

    user_ids, password = scrape_user_credentials(URL)
    write_to_excel(user_ids, password, OUTPUT_FILE)
