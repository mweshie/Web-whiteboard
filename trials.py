from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Path to your ChromeDriver executable
chromedriver_path = 'C:\Webdrivers\chromedriver-win64\chromedriver-win64\chromedriver.exe'

# Set up the ChromeDriver service
service = Service(chromedriver_path)

# Initialize the Chrome WebDriver
driver = webdriver.Chrome(service=service)

# Function to log in
def login_aad_user(username, password):
    driver.get("https://www.office.com/")
    sign_in_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.LINK_TEXT, "Sign in"))
    )
    sign_in_button.click()
    
    email_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "loginfmt"))
    )
    email_input.send_keys(username)
    email_input.send_keys(Keys.RETURN)
    
    time.sleep(5)  # Adjust based on actual page loading times
    
    password_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "passwd"))
    )
    password_input.send_keys(password)
    password_input.send_keys(Keys.RETURN)
    
    time.sleep(2)  # Adjust based on actual page loading times

    # Authentication part
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "signInAnotherWay")))
        driver.find_element(By.ID, "signInAnotherWay").click()
    except Exception as e:
        print(f"Error during selecting another way to sign in: {e}")
        return

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".row:nth-child(3) .text-left > div")))
        driver.find_element(By.CSS_SELECTOR, ".row:nth-child(3) .text-left > div").click()
    except Exception as e:
        print(f"Error during selecting the third option: {e}")
        return

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "idTxtBx_SAOTCC_OTC")))
        driver.find_element(By.ID, "idTxtBx_SAOTCC_OTC").click()
    except Exception as e:
        print(f"Error during entering authentication code: {e}")
        return

    # Wait for the authentication code and enter manually
    time.sleep(30)  # Adjust sleep time based on how long it takes to receive the 2FA code

    driver.find_element(By.ID, "idSubmit_SAOTCC_Continue").click()

try:
    # Step 1: Sign in with an AAD User
    login_aad_user('nestorw@M365x82407187.onmicrosoft.com', 'e~Gb#C^22V(bk32E')

    # Step 2: Search for Whiteboard on the search bar
    search_bar = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search']"))
    )
    search_bar.send_keys("Whiteboard")
    search_bar.send_keys(Keys.RETURN)
    
    # Step 3: Open a new board
    whiteboard_link = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Whiteboard"))
    )
    whiteboard_link.click()
    
    time.sleep(5)  # Adjust based on actual page loading times

    # Step 4: Click on the settings icon on the top right
    settings_icon = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "button[title='Settings']"))
    )
    settings_icon.click()
    
    # Step 5: Navigate to the backstage by selecting 'Privacy and Settings'
    privacy_settings_link = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Privacy and Settings"))
    )
    privacy_settings_link.click()
    
    time.sleep(3)  # Adjust based on actual page loading times
    
    # Step 6: Navigate to the “Optional Connected Experiences” check box
    optional_connected_experiences_checkbox = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='checkbox'][name='optionalConnectedExperiences']"))
    )

    # Step 7: Verify that the “Optional Connected Experiences” switch is visible
    assert optional_connected_experiences_checkbox.is_displayed(), "Optional Connected Experiences switch not found."

    print("Test passed: Optional Connected Experiences switch is visible.")

except Exception as e:
    print(f"Test failed: {e}")

finally:
    # Close the browser
    driver.quit()
