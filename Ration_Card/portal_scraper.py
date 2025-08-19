import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, UnexpectedAlertPresentException
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class RajasthanFoodPortalScraper:
    def __init__(self, headless=True):
        """Initialize the scraper with Chrome WebDriver"""
        self.chrome_options = Options()
        if headless:
            self.chrome_options.add_argument("--headless")
        self.chrome_options.add_argument("--no-sandbox")
        self.chrome_options.add_argument("--disable-dev-shm-usage")
        self.chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        self.chrome_options.add_argument("--disable-web-security")
        self.chrome_options.add_argument("--allow-running-insecure-content")
        self.chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.chrome_options.add_experimental_option('useAutomationExtension', False)
        
        self.driver = None
        self.portal_url = "https://food.rajasthan.gov.in/Form_Status.aspx"
    
    def start_driver(self):
        """Start the Chrome WebDriver"""
        try:
            self.driver = webdriver.Chrome(options=self.chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.driver.implicitly_wait(10)
            logger.info("WebDriver started successfully")
        except Exception as e:
            logger.error(f"Failed to start WebDriver: {str(e)}")
            raise
    
    def handle_alert(self):
        """Handle any unexpected alerts"""
        try:
            alert = self.driver.switch_to.alert
            alert_text = alert.text
            logger.warning(f"Alert detected: {alert_text}")
            alert.accept()
            return alert_text
        except:
            return None
    
    def search_ration_card(self, ration_card_number):
        """Search for ration card details"""
        try:
            logger.info(f"Searching for ration card: {ration_card_number}")
            
            # Navigate to the portal
            self.driver.get(self.portal_url)
            time.sleep(3)
            
            # Handle any initial alerts
            self.handle_alert()
            
            # Find input field
            input_element = None
            input_selectors = [
                "input[id*='txt']",
                "input[name*='txt']", 
                "input[type='text']"
            ]
            
            for selector in input_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for elem in elements:
                        if elem.is_displayed() and elem.is_enabled():
                            input_element = elem
                            logger.info(f"Found input field with selector: {selector}")
                            break
                    if input_element:
                        break
                except:
                    continue
            
            if not input_element:
                return {"error": "Could not find input field"}
            
            # Clear and enter ration card number
            try:
                input_element.clear()
                time.sleep(0.5)
                input_element.click()
                time.sleep(0.5)
                input_element.send_keys(Keys.CONTROL + "a")
                input_element.send_keys(Keys.DELETE)
                time.sleep(0.5)
                input_element.send_keys(ration_card_number)
                time.sleep(1)
                
                entered_value = input_element.get_attribute('value')
                logger.info(f"Entered value: '{entered_value}'")
                
                if entered_value != ration_card_number:
                    self.driver.execute_script(f"arguments[0].value = '{ration_card_number}';", input_element)
                    self.driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", input_element)
                    self.driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", input_element)
                    time.sleep(1)
                
            except Exception as e:
                logger.error(f"Error entering ration card number: {str(e)}")
                return {"error": f"Failed to enter ration card number: {str(e)}"}
            
            # Find and click submit button
            submit_button = None
            submit_selectors = [
                "input[type='submit']",
                "button[type='submit']",
                "input[value*='Search']",
                "input[id*='btn']"
            ]
            
            for selector in submit_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for elem in elements:
                        if elem.is_displayed() and elem.is_enabled():
                            submit_button = elem
                            logger.info(f"Found submit button with selector: {selector}")
                            break
                    if submit_button:
                        break
                except:
                    continue
            
            if not submit_button:
                return {"error": "Submit button not found"}
            
            # Click submit
            try:
                self.driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
                time.sleep(1)
                submit_button.click()
                logger.info("Clicked submit button")
            except Exception as e:
                try:
                    self.driver.execute_script("arguments[0].click();", submit_button)
                    logger.info("Clicked submit button with JavaScript")
                except:
                    return {"error": f"Failed to click submit button: {str(e)}"}
            
            # Wait for results
            time.sleep(4)
            
            # Handle any alerts
            alert_text = self.handle_alert()
            if alert_text and "Please Enter" in alert_text:
                return {"error": f"Form validation failed: {alert_text}"}
            
            # Extract results
            return self.extract_results(ration_card_number)
                
        except Exception as e:
            logger.error(f"Error during search: {str(e)}")
            return {"error": f"Search failed: {str(e)}"}
    
    def extract_results(self, ration_card_number):
        """Extract ration card details with improved parsing"""
        try:
            time.sleep(2)
            
            # Handle any alerts during extraction
            alert_text = self.handle_alert()
            if alert_text:
                return {"error": f"Alert during extraction: {alert_text}"}
            
            # Initialize data structure
            data = {
                "ration_card_number": ration_card_number,
                "search_timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                "status": "searched"
            }
            
            # Extract all tables
            tables = self.driver.find_elements(By.TAG_NAME, "table")
            if tables:
                logger.info(f"Found {len(tables)} tables on results page")
                for i, table in enumerate(tables):
                    table_text = table.text.strip()
                    if table_text:
                        data[f"table_{i}_content"] = table_text
                        logger.info(f"Table {i} content: {table_text[:100]}...")
            
            # Try to find the main data table with ration card information
            main_data_found = False
            for key, content in data.items():
                if key.startswith('table_') and content:
                    # Look for ration card data patterns
                    if ('प्राधिकृत अधिकारी' in content or 'Officer' in content) and ration_card_number in content:
                        main_data_found = True
                        data['main_table'] = content
                        logger.info(f"Found main data in {key}")
                        break
                    # Also check for pattern with form numbers and token numbers
                    elif any(len(word) >= 8 and word.isdigit() for word in content.split()):
                        main_data_found = True
                        data['main_table'] = content
                        logger.info(f"Found numeric data in {key}")
                        break
            
            # Get full page text as backup
            try:
                body = self.driver.find_element(By.TAG_NAME, "body")
                body_text = body.text.strip()
                if body_text:
                    data["full_page_text"] = body_text
            except:
                pass
            
            # Check if we have meaningful results
            if main_data_found or any('प्राधिकृत' in str(value) for value in data.values()):
                logger.info("✅ Found ration card data successfully")
            else:
                # Check for common "not found" messages
                page_text = data.get('full_page_text', '').lower()
                if any(phrase in page_text for phrase in ['no record', 'not found', 'no data']):
                    data["error"] = "No records found for this ration card number"
                elif len(page_text) < 200:  # Very short response might indicate an error
                    data["error"] = "Minimal response from portal - possible error"
            
            logger.info("Data extraction completed")
            return data
            
        except Exception as e:
            logger.error(f"Error extracting results: {str(e)}")
            return {
                "error": f"Failed to extract results: {str(e)}",
                "ration_card_number": ration_card_number
            }
    
    def close(self):
        """Close the WebDriver"""
        if self.driver:
            self.driver.quit()
            logger.info("WebDriver closed")