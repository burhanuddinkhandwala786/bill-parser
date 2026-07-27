import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

STREAMLIT_URL = "https://YOUR-APP-NAME.streamlit.app"  # <--- Put your app URL here

def wake_app():
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    
    driver = webdriver.Chrome(options=options)
    print(f"Navigating to {STREAMLIT_URL}...")
    driver.get(STREAMLIT_URL)
    
    wait = WebDriverWait(driver, 15)
    
    try:
        # Check if the "Yes, get this app back up!" button exists (App is asleep)
        wake_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Yes, get this app back up')]"))
        )
        print("App is asleep! Clicking wake button...")
        wake_button.click()
        print("Wake button clicked successfully. App is starting up!")
    except Exception:
        print("No sleep button detected. App is already awake and active!")
    finally:
        driver.quit()

if __name__ == "__main__":
    wake_app()
