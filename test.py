import unittest
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.chrome.service import Service
import time


class RoamingNotification(unittest.TestCase):

    def setUp(self):
        """Set up Chrome driver."""
        try:
            # Initialize Chrome driver
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument("--incognito")
            self.chrome_driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
            self.chrome_driver.set_window_size(945, 1012)
        except Exception as e:
            print(f"Error during Chrome setup: {e}")
            self.chrome_driver = None

    def tearDown(self):
        """Tear down the driver after tests."""
        if self.chrome_driver:
            self.chrome_driver.quit()

    def test_login_and_authenticate(self):
        """Test login and authenticate the user in the specified driver."""
        if self.chrome_driver is None:
            return
        try:
            # Navigate to the login page
            self.chrome_driver.get("https://office.com")

            # Click on the Sign in button
            sign_in_button = WebDriverWait(self.chrome_driver, 10).until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Sign in"))
            )
            sign_in_button.click()

            # Retry logic for finding and interacting with elements to avoid stale references
            def safe_find_element(driver, by, value, retries=3):
                for _ in range(retries):
                    try:
                        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((by, value)))
                        return element
                    except StaleElementReferenceException:
                        print("Stale element, retrying...")
                raise Exception(f"Element {value} not stable after {retries} retries")

            # Enter email and proceed
            email_input = safe_find_element(self.chrome_driver, By.NAME, "loginfmt")
            email_input.send_keys("AllanD@M365x74150002.OnMicrosoft.com")
            email_input.send_keys(Keys.ENTER)
            time.sleep(5)
            
            # Wait for the password input to appear
            password_input = safe_find_element(self.chrome_driver, By.NAME, "passwd")
            password_input.send_keys("Kenya@2030")

            # Click the Sign in button after entering the password
            try:
                sign_in_button = WebDriverWait(self.chrome_driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "idSIButton9"))
                )
                sign_in_button.click()
            except TimeoutException:
                # Try using a different method to find the button, if needed
                sign_in_button = WebDriverWait(self.chrome_driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='submit']"))
                )
                sign_in_button.click()

            time.sleep(20)

            stay_signed_in_button = WebDriverWait(self.chrome_driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "idSIButton9"))
                )
            stay_signed_in_button.click()

            time.sleep(10)

            select_office_python_apps(self.chrome_driver)
            self.chrome_driver.switch_to.window(self.chrome_driver.window_handles[1])
            navigate_to_backstage(self.chrome_driver)
            adjust_privacy_settings(self.chrome_driver)
            refresh_settings(self.chrome_driver)
            time.sleep(5)
            navigate_to_backstage(self.chrome_driver)

        except TimeoutException:
            print("Element not found or timed out.")
        except Exception as e:
            print(f"Error during login and authentication: {e}")

        time.sleep(60)


# Additional feature functions
def select_office_python_apps(driver):
    apps_btn = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='26706956-2AAE-4A22-8960-09B98C35B28C']"))
    )
    apps_btn.click()

    word_app = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//*[@id='core-apps']/div[1]/div[3]"))
    )
    word_app.click()

    blank_doc_btn = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "[aria-label='Blank document']"))
    )
    blank_doc_btn.click()

    time.sleep(10)

def navigate_to_backstage(driver):
    iframe = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "WacFrame_Word_0"))
    )
    driver.switch_to.frame(iframe)

    file_btn = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "FileMenuLauncherContainer"))
    )
    file_btn.click()

    info_btn = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "Info"))
    )

    info_btn.click()

    privacy_settings_btn = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, "Privacy settings"))
    )
    privacy_settings_btn.click()

    time.sleep(2)

def adjust_privacy_settings(driver):
    toggle_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//*[@id='OptionalConnectedExperiencesDialogCheckbox']/div/label/div"))
    )
    toggle_button.click()

    ok_btn = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "DialogActionButton"))
    )
    ok_btn.click()

    time.sleep(2)

def refresh_settings(driver):
    refresh_btn = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "DialogActionButton"))
    )
    refresh_btn.click()


if __name__ == "__main__":
    unittest.main()
