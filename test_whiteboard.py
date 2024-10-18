import unittest
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.chrome.service import Service
import time

class RoamingNotification(unittest.TestCase):

    def setUp(self):
        """Set up Chrome and Edge drivers."""
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
        """Tear down the drivers after tests."""
        if self.chrome_driver:
            self.chrome_driver.quit()

    def test_login_and_authenticate(self):
        """Test login and authenticate the user in the specified driver."""
        if self.chrome_driver is None:
            return
        try:
            # Navigate to the login page
            self.chrome_driver.get("https://office.com")

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
            email_input = safe_find_element(self.chrome_driver, By.ID, "i0116")
            email_input.send_keys("AllanD@M365x74150002.OnMicrosoft.com")
            email_input.send_keys(Keys.ENTER)

            # Enter password and proceed
            password_input = safe_find_element(self.chrome_driver, By.ID, "i0118")
            password_input.send_keys("Kenya@2030")
            password_input.send_keys(Keys.ENTER)

            # Select "Sign in another way"
            sign_in_another_way_button = safe_find_element(self.chrome_driver, By.ID, "signInAnotherWay")
            sign_in_another_way_button.click()

            # Select the third option
            third_option = safe_find_element(self.chrome_driver, By.CSS_SELECTOR, ".row:nth-child(3) .text-left > div")
            third_option.click()

            # Click on the authentication code field
            auth_code_field = safe_find_element(self.chrome_driver, By.ID, "idTxtBx_SAOTCC_OTC")
            auth_code_field.click()

            # Wait for the user to manually enter the authentication code
            time.sleep(30)  # Adjust sleep time based on how long it takes to receive the 2FA code

            # Click continue after entering the authentication code
            continue_button = safe_find_element(self.chrome_driver, By.ID, "idSubmit_SAOTCC_Continue")
            continue_button.click()
            
        except TimeoutException:
            print(f"Element not found or timed out.")
        except Exception as e:
            print(f"Error during login and authentication: {e}")

    def test_whiteboard_search(self):
        # Step 2: Search for Whiteboard on the search bar
        search_bar = WebDriverWait(self.chrome_driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='search']"))
        )
        search_bar.send_keys("Whiteboard")
        search_bar.send_keys(Keys.RETURN)
        
        # Step 3: Open a new board
        whiteboard_link = WebDriverWait(self.chrome_driver, 10).until(
            EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Whiteboard"))
        )
        whiteboard_link.click()
        
        time.sleep(5)  # Adjust based on actual page loading times

    
if __name__ == "__main__":
    unittest.main()
