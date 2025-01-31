from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException
import time
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv
import os

load_dotenv()

ratte, sheet_id = os.getenv("RATTE"), os.getenv("SHEET_ID")
SCOPES, SERVICE_ACCOUNT_FILE = ['https://www.googleapis.com/auth/spreadsheets'], 'google_auth.json'
credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
gs_service = build("sheets", "v4", credentials=credentials)

firefox_options = Options()
firefox_options.add_argument("--headless")
driver = webdriver.Firefox(service=Service(), options=firefox_options)


# Update Google Sheet
def update_values(range_name, value_input_option, values, fish):
    body = {'values': values}
    result = gs_service.spreadsheets().values().update(
        spreadsheetId=sheet_id, range=range_name, valueInputOption=value_input_option, body=body).execute()
    if result:
        print(f"{result.get('updatedCells')} cells of {fish.capitalize()} updated.")
    return result


# Open page and retrieve driver
def open_page(page):
    try:
        url = f"https://www.seaofthieves.com/profile/reputation/{page}"
        driver.get(url)
        driver.add_cookie({"name": "rat", "value": ratte})
        driver.get(url)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CLASS_NAME, "emblem-item__progress-text")))
        WebDriverWait(driver, 5)
        return driver.page_source
    except WebDriverException as e:
        print(f"Error opening page {page}: {e}.\nProb incorrect RATTE cookie.\n")
        return None
    except Exception as e:
        print(f"Unexpected error opening page {page}: {e}")
        return None


def get_data(fish, page_source):
    try:
        if not page_source:
            return None
        # Process the page source and return it
        return str(page_source.encode("utf-8"))
    except WebDriverException as e:
        print(f"Error scraping {fish}: {e}. Restarting WebDriver...")
        driver.quit()  # Ensure the old driver is quit before restarting
        return get_data(fish, page_source)  # Retry with an increased timeout
    except Exception as e:
        print(f"Error scraping {fish}: {e}")
        return None


def find_achievements(src, indices):
    result = []
    while "<div class=\"emblem-item__progress-text\">" in src:
        divindex = src.find("<div class=\"emblem-item__progress-text\">")
        src = src[divindex + 40:]
        result.append(src[:src.find("/")])
    return [result[i] for i in indices if i < len(result)]


def get_list(fish, indices, page_source):
    print(f"Updating {fish.capitalize()}")
    fishcaught = find_achievements(get_data(fish, page_source), indices)
    if not fishcaught:
        print(f"Retrying {fish} with a longer timeout.")
        fishcaught = find_achievements(get_data(fish, page_source), indices)
        if not fishcaught:
            print(f"Failed to update {fish}. Skipping.")
            return []
    return [[fish] for fish in fishcaught]


update_functions = {
    "splashtails": ("E5:E9", [0, 1, 2, 3, 4], "HuntersCall", []),
    "wildsplashes": ("E13:E17", [0, 1, 2, 3, 4], "HuntersCall", []),
    "pondies": ("E21:E25", [0, 1, 2, 3, 4], "HuntersCall", []),
    "wreckers": ("W5:W9", [0, 1, 2, 3, 4], "HuntersCall", []),
    "cooking": ("AE6:BAE11", [0, 1, 2, 3, 4, 5], "HuntersCall", []),
    "plentifins": ("K5:K9", [0, 1, 2, 3, 4], "HuntersCall", []),
    "devilfishes": ("K13:K17", [0, 1, 2, 3, 4], "HuntersCall", []),
    "battlegills": ("K21:K25", [0, 1, 2, 3, 4], "HuntersCall", []),
    "ancientscales": ("Q5:Q9", [0, 1, 2, 3, 4], "HuntersCall", []),
    "islehoppers": ("Q13:Q17", [0, 1, 2, 3, 4], "HuntersCall", []),
    "stormfishes": ("Q21:Q25", [0, 1, 2, 3, 4], "HuntersCall", []),
    "merrick's-accolades": ("AE14", [0], "HuntersCall", []),
    "shrouded-spoils": ("AE19:AE24", [2, 3, 4, 5, 6, 7], "BilgeRats", []),
}

for fish_name, (cell_range, indices, page, accolades) in update_functions.items():
    page_source = open_page(f"{page}/{fish_name}")
    fishlist = get_list(fish_name, indices, page_source)
    if fishlist:
        update_values(cell_range, "USER_ENTERED", fishlist, fish_name)
