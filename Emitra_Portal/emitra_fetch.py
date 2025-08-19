from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import gspread
from google.oauth2.service_account import Credentials
import time
import logging
import sys
import os

# Fix Windows encoding issues
if sys.platform == "win32":
    os.environ['PYTHONIOENCODING'] = 'utf-8'

# Simple logging without emojis
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('emitra_automation.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

class EmitraCleanAutomation:
    def __init__(self):
        # Google Sheets setup
        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
        self.client = gspread.authorize(creds)
        self.sheet = self.client.open('Automation sheet').worksheet('Emitra')
        
        # Chrome setup - fully headless
        chrome_options = Options()
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-plugins")
        chrome_options.add_argument("--disable-images")
        chrome_options.add_argument("--disable-logging")
        chrome_options.add_argument("--log-level=3")
        chrome_options.add_argument("--silent")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # User agent
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.wait = WebDriverWait(self.driver, 30)
        
        # Stealth configuration
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        logging.info("Emitra Automation Started Successfully")
    
    def wait_for_angular_load(self):
        """Wait for Angular page to load with better detection"""
        logging.info("Waiting for Angular page to load...")
        
        try:
            # Wait for basic page structure
            self.wait.until(EC.presence_of_element_located((By.TAG_NAME, "app-root")))
            
            # Wait for verification card
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.card.verification-transaction")))
            
            # Wait for radio buttons to be present
            self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "mat-radio-button")))
            
            # Wait for input field to be present
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.form-control")))
            
            # Additional wait for JavaScript to complete
            time.sleep(5)
            
            logging.info("Angular components loaded successfully")
            return True
            
        except Exception as e:
            logging.error(f"Failed to load Angular page: {str(e)}")
            return False
    
    def select_receipt_number_option(self):
        """Select Receipt Number radio button with better error handling"""
        logging.info("Selecting Receipt Number option...")
        
        try:
            # First try the specific Receipt Number radio button
            receipt_radio = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "mat-radio-button[value='2']"))
            )
            
            # Click using JavaScript if regular click fails
            try:
                receipt_radio.click()
            except:
                self.driver.execute_script("arguments[0].click();", receipt_radio)
            
            time.sleep(3)
            logging.info("Receipt Number option selected")
            return True
            
        except Exception as e:
            logging.error(f"Failed to select Receipt Number option: {str(e)}")
            
            # Fallback: Try to find any radio button related to receipts
            try:
                all_radios = self.driver.find_elements(By.CSS_SELECTOR, "mat-radio-button")
                for radio in all_radios:
                    if "Receipt" in radio.text:
                        self.driver.execute_script("arguments[0].click();", radio)
                        time.sleep(3)
                        logging.info("Receipt option selected via fallback")
                        return True
            except:
                pass
            
            return False
    
    def enter_receipt_number(self, receipt_number):
        """Enter receipt number with multiple input field strategies"""
        logging.info(f"Entering receipt number: {receipt_number}")
        
        # Try multiple input field selectors
        input_selectors = [
            "input.form-control[placeholder*='12/16 Digit Number']",
            "input.form-control",
            "input[type='search']",
            "input[placeholder*='Digit']",
            "input[maxlength='16']",
            ".input-group input"
        ]
        
        for selector in input_selectors:
            try:
                input_field = self.wait.until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
                
                # Clear and enter receipt number
                input_field.clear()
                time.sleep(1)
                input_field.send_keys(receipt_number)
                
                # Verify the input was entered
                if input_field.get_attribute('value') == receipt_number:
                    logging.info(f"Successfully entered receipt number: {receipt_number}")
                    return True
                    
            except Exception as e:
                logging.warning(f"Input selector {selector} failed: {str(e)}")
                continue
        
        logging.error("Failed to find or fill receipt input field")
        return False
    
    def click_search_button(self):
        """Click search button with multiple strategies"""
        logging.info("Clicking Search button...")
        
        try:
            # Wait for search button to be enabled
            search_button = self.wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-outline-primary.searchBtnnew"))
            )
            
            # Try regular click first
            try:
                search_button.click()
            except:
                # Fallback to JavaScript click
                self.driver.execute_script("arguments[0].click();", search_button)
            
            logging.info("Search button clicked, waiting for results...")
            time.sleep(8)  # Longer wait for results
            return True
            
        except Exception as e:
            logging.error(f"Failed to click search button: {str(e)}")
            
            # Try alternative search button selectors
            search_selectors = [
                "button[type='button']:contains('Search')",
                "button.searchBtnnew",
                ".btn:contains('Search')"
            ]
            
            for selector in search_selectors:
                try:
                    if ":contains(" in selector:
                        continue  # Skip CSS :contains() as it's not supported
                    button = self.driver.find_element(By.CSS_SELECTOR, selector)
                    self.driver.execute_script("arguments[0].click();", button)
                    time.sleep(8)
                    logging.info("Search clicked via fallback")
                    return True
                except:
                    continue
            
            return False
    
    def _clean_service_name(self, raw_text):
        """Clean and format the service name by removing unwanted prefixes and suffixes"""
        try:
            if not raw_text or len(raw_text.strip()) < 3:
                return None
            
            text = raw_text.strip()
            
            # Remove common prefixes (case insensitive)
            prefixes_to_remove = [
                "Service :",
                "Service:",
                "Service Name :",
                "Service Name:",
                "Service Type :",
                "Service Type:",
                "Application :",
                "Application:",
                "Form :",
                "Form:",
                "Certificate :",
                "Certificate:",
                "Name :",
                "Name:"
            ]
            
            for prefix in prefixes_to_remove:
                if text.lower().startswith(prefix.lower()):
                    text = text[len(prefix):].strip()
                    break
            
            # Remove common suffixes
            suffixes_to_remove = [
                "- Click for more details",
                "- View More",
                "- More Info",
                "- Details",
                "Click here",
                "View More"
            ]
            
            for suffix in suffixes_to_remove:
                if text.lower().endswith(suffix.lower()):
                    text = text[:-len(suffix)].strip()
                    break
            
            # Clean up extra spaces and special characters
            text = ' '.join(text.split())  # Remove extra whitespace
            text = text.replace('\n', ' ').replace('\r', ' ')  # Remove line breaks
            
            # Remove leading/trailing quotes or brackets if present
            text = text.strip('"\'()[]{}')
            
            # Skip if it's just common words that aren't service names
            skip_words = ['search', 'result', 'receipt', 'number', 'click', 'view', 'more', 'date', 'time', 'status', 'here', 'details']
            if text.lower() in skip_words:
                return None
            
            # Must have meaningful content
            if len(text) < 5 or len(text) > 200:
                return None
            
            # Should contain service-related keywords or be substantial text
            service_keywords = ['certificate', 'registration', 'license', 'verification', 'application', 'form', 'permit', 'approval']
            if len(text) < 20 and not any(keyword in text.lower() for keyword in service_keywords):
                return None
            
            return text
            
        except Exception as e:
            logging.debug(f"Error cleaning service name '{raw_text}': {e}")
            return None
    
    def extract_service_name(self):
        """Extract service name from the search results page BEFORE clicking View More"""
        logging.info("Extracting service name from search results...")
        
        try:
            time.sleep(3)  # Wait for search results to load
            
            # Strategy 1: Look for service name in common locations
            service_selectors = [
                "//div[contains(@class, 'service')]//text()[normalize-space()]",
                "//span[contains(text(), 'Service')]/..//text()[normalize-space()]",
                "//div[contains(text(), 'Service')]//text()[normalize-space()]",
                "//td[contains(text(), 'Service')]/following-sibling::td//text()[normalize-space()]",
                "//label[contains(text(), 'Service')]/..//text()[normalize-space()]",
                "//div[contains(@class, 'card-body')]//strong//text()[normalize-space()]",
                "//div[contains(@class, 'result')]//h5//text()[normalize-space()]",
                "//div[contains(@class, 'result')]//h6//text()[normalize-space()]"
            ]
            
            for selector in service_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for element in elements:
                        text = element.text.strip() if hasattr(element, 'text') else str(element).strip()
                        if text and len(text) > 3 and len(text) < 200:
                            # Clean the service name - remove prefixes
                            cleaned_text = self._clean_service_name(text)
                            if cleaned_text:
                                logging.info(f"Found service name: {cleaned_text}")
                                return cleaned_text
                except Exception as e:
                    logging.debug(f"Service selector {selector} failed: {e}")
                    continue
            
            # Strategy 2: Look in the main content area
            try:
                content_areas = self.driver.find_elements(By.CSS_SELECTOR, "div.card-body, .result-content, .search-result")
                for area in content_areas:
                    text_elements = area.find_elements(By.XPATH, ".//*[text()]")
                    for elem in text_elements:
                        text = elem.text.strip()
                        if text and 10 <= len(text) <= 150:
                            # Look for service-like patterns
                            if any(keyword in text.lower() for keyword in ['certificate', 'registration', 'license', 'verification', 'application']):
                                cleaned_text = self._clean_service_name(text)
                                if cleaned_text:
                                    logging.info(f"Found service name via content area: {cleaned_text}")
                                    return cleaned_text
            except:
                pass
            
            # Strategy 3: Look for any meaningful text in results
            try:
                # Get all visible text from the results area
                results_div = self.driver.find_element(By.CSS_SELECTOR, "div[style*='background']")
                all_text = results_div.text.strip()
                
                if all_text:
                    lines = [line.strip() for line in all_text.split('\n') if line.strip()]
                    for line in lines:
                        if 10 <= len(line) <= 200:
                            cleaned_text = self._clean_service_name(line)
                            if cleaned_text and not any(skip in cleaned_text.lower() for skip in ['view more', 'receipt', 'search', 'click', 'here']):
                                logging.info(f"Found service name via text analysis: {cleaned_text}")
                                return cleaned_text
            except:
                pass
            
            logging.warning("Could not extract service name from search results")
            return "SERVICE NAME NOT FOUND"
            
        except Exception as e:
            logging.error(f"Error extracting service name: {str(e)}")
            return "SERVICE EXTRACTION ERROR"
    
    def click_view_more(self):
        """Click VIEW MORE with enhanced detection"""
        logging.info("Looking for VIEW MORE link...")
        
        try:
            # Wait for results section to appear
            self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[style*='background']")))
            
            # Multiple strategies for VIEW MORE
            view_more_selectors = [
                "//small[text()='VIEW MORE']",
                "//small[contains(text(), 'VIEW MORE')]",
                "//div[contains(@class, 'link')]//small",
                "//div[contains(@class, 'text-end')]//small",
                "//small"
            ]
            
            for selector in view_more_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for element in elements:
                        if "VIEW MORE" in element.text.upper():
                            # Try regular click
                            try:
                                element.click()
                            except:
                                # Try JavaScript click
                                self.driver.execute_script("arguments[0].click();", element)
                            
                            time.sleep(5)
                            logging.info("VIEW MORE clicked successfully")
                            return True
                except:
                    continue
            
            logging.warning("VIEW MORE not found, proceeding anyway")
            return True  # Continue even if not found
            
        except Exception as e:
            logging.error(f"Error with VIEW MORE: {str(e)}")
            return True  # Continue anyway
    
    def click_lifecycle_tab(self):
        """Click Life Cycle tab if available"""
        logging.info("Looking for Life Cycle tab...")
        
        try:
            time.sleep(5)  # Wait for content to expand
            
            lifecycle_selectors = [
                "//button[contains(text(), 'Life cycle')]",
                "//button[contains(text(), 'Life Cycle')]",
                "//a[contains(text(), 'Life cycle')]",
                "//span[contains(text(), 'Life cycle')]",
                "//div[contains(text(), 'Life cycle')]"
            ]
            
            for selector in lifecycle_selectors:
                try:
                    element = self.driver.find_element(By.XPATH, selector)
                    if element.is_displayed():
                        try:
                            element.click()
                        except:
                            self.driver.execute_script("arguments[0].click();", element)
                        
                        time.sleep(3)
                        logging.info("Life Cycle tab clicked")
                        return True
                except:
                    continue
            
            logging.info("Life Cycle tab not found, data might already be visible")
            return True
            
        except Exception as e:
            logging.error(f"Error with Life Cycle tab: {str(e)}")
            return True
    
    def extract_lifecycle_data(self):
        """Extract data with comprehensive strategies"""
        logging.info("Extracting lifecycle data...")
        
        try:
            time.sleep(5)  # Wait for data to load
            
            # Strategy 1: Table data extraction
            table_selectors = [
                "//table//tbody//tr[td]",
                "//table//tr[td]",
                "//div[contains(@class, 'table')]//tr[td]",
                "//mat-table//mat-row",
                "//tr[position()>1 and td]"
            ]
            
            for selector in table_selectors:
                try:
                    rows = self.driver.find_elements(By.XPATH, selector)
                    if len(rows) > 0:
                        logging.info(f"Found {len(rows)} table rows")
                        
                        # Get the last row (most recent)
                        last_row = rows[-1]
                        cells = last_row.find_elements(By.XPATH, ".//td | .//mat-cell")
                        
                        if len(cells) >= 3:
                            data = []
                            for cell in cells:
                                text = cell.text.strip()
                                if text:
                                    data.append(text)
                            
                            if data:
                                # Ensure exactly 6 columns (keeping original 6 lifecycle columns)
                                while len(data) < 6:
                                    data.append('')
                                if len(data) > 6:
                                    data = data[:6]
                                
                                logging.info(f"Successfully extracted table data: {data[0] if data[0] else 'Empty'}")
                                return data
                except:
                    continue
            
            # Strategy 2: Any structured data
            try:
                all_text_elements = self.driver.find_elements(By.XPATH, "//*[text()]")
                data_keywords = ['date', 'time', 'status', 'officer', 'location', 'remark']
                extracted_data = []
                
                for element in all_text_elements:
                    text = element.text.strip().lower()
                    if any(keyword in text for keyword in data_keywords):
                        parent_text = element.find_element(By.XPATH, "..").text.strip()
                        if parent_text and len(parent_text) > 5:
                            extracted_data.append(parent_text)
                            if len(extracted_data) >= 6:
                                break
                
                if extracted_data:
                    while len(extracted_data) < 6:
                        extracted_data.append('')
                    logging.info("Found structured data elements")
                    return extracted_data[:6]
                    
            except:
                pass
            
            # Strategy 3: Look for any meaningful data
            try:
                page_text = self.driver.page_source
                if "success" in page_text.lower() or "completed" in page_text.lower():
                    return ["DATA FOUND BUT NOT STRUCTURED"] * 6
                elif "not found" in page_text.lower() or "invalid" in page_text.lower():
                    return ["RECEIPT NOT FOUND"] * 6
            except:
                pass
            
            logging.warning("No lifecycle data found")
            return ["NO DATA AVAILABLE"] * 6
            
        except Exception as e:
            logging.error(f"Error extracting data: {str(e)}")
            return ["EXTRACTION ERROR"] * 6
    
    def process_single_receipt(self, receipt_number, row_index):
        """Process one receipt with comprehensive error handling"""
        logging.info(f"Processing receipt {receipt_number} (Row {row_index})")
        
        try:
            # Load page
            self.driver.get("https://emitra.rajasthan.gov.in/emitra/home")
            
            # Wait for page load
            if not self.wait_for_angular_load():
                return "PAGE LOAD FAILED", ["PAGE LOAD FAILED"] * 6
            
            # Handle popup
            try:
                popup = self.driver.find_element(By.CSS_SELECTOR, "button.p-dialog-header-close")
                popup.click()
                time.sleep(2)
                logging.info("Popup closed")
            except:
                logging.info("No popup found")
            
            # Execute the flow
            if not self.select_receipt_number_option():
                return "RECEIPT SELECTION FAILED", ["RECEIPT SELECTION FAILED"] * 6
            
            if not self.enter_receipt_number(receipt_number):
                return "INPUT FAILED", ["INPUT FAILED"] * 6
            
            if not self.click_search_button():
                return "SEARCH FAILED", ["SEARCH FAILED"] * 6
            
            # NEW: Extract service name BEFORE clicking View More
            service_name = self.extract_service_name()
            
            if not self.click_view_more():
                return service_name, ["VIEW MORE FAILED"] * 6
            
            self.click_lifecycle_tab()
            lifecycle_data = self.extract_lifecycle_data()
            
            logging.info(f"Processing complete for {receipt_number}: Service='{service_name}', Lifecycle='{lifecycle_data[0] if lifecycle_data else 'No result'}'")
            return service_name, lifecycle_data
            
        except Exception as e:
            logging.error(f"Error processing {receipt_number}: {str(e)}")
            return "PROCESSING ERROR", ["PROCESSING ERROR"] * 6
    
    def run_automation(self):
        """Main automation runner"""
        start_time = time.time()
        logging.info("STARTING EMITRA AUTOMATION WITH CLEAN SERVICE NAME EXTRACTION")
        logging.info("Current Date and Time (UTC): 2025-08-18 10:36:44")
        logging.info("Current User's Login: deepanshudagdi")
        logging.info("=" * 50)
        
        try:
            # Get receipts from sheet
            receipt_numbers = self.sheet.col_values(1)[1:]  # Skip header
            valid_receipts = [(r.strip(), idx+2) for idx, r in enumerate(receipt_numbers) if r.strip()]
            total = len(valid_receipts)
            successful = 0
            failed = 0
            
            logging.info(f"Found {total} receipts to process")
            
            for i, (receipt_number, row_index) in enumerate(valid_receipts, 1):
                logging.info(f"[{i}/{total}] Processing: {receipt_number}")
                
                service_name, lifecycle_data = self.process_single_receipt(receipt_number, row_index)
                
                # Combine service name with lifecycle data
                # Service name goes to column B, lifecycle data goes to columns C-H
                combined_result = [service_name] + lifecycle_data
                
                # Update Google Sheet - now writing to columns B through H (7 columns total)
                try:
                    self.sheet.update(range_name=f'B{row_index}:H{row_index}', values=[combined_result])
                    
                    # Check if successful
                    if (service_name and 
                        "ERROR" not in service_name.upper() and 
                        "FAILED" not in service_name.upper() and
                        lifecycle_data[0] and
                        "ERROR" not in lifecycle_data[0].upper() and 
                        "FAILED" not in lifecycle_data[0].upper() and
                        "NO DATA" not in lifecycle_data[0].upper()):
                        successful += 1
                        logging.info(f"[{i}/{total}] SUCCESS: {receipt_number} - Service: {service_name[:50]}...")
                    else:
                        failed += 1
                        logging.warning(f"[{i}/{total}] PARTIAL: {receipt_number} - Service: {service_name}, Lifecycle: {lifecycle_data[0]}")
                        
                except Exception as e:
                    failed += 1
                    logging.error(f"[{i}/{total}] SHEET ERROR: {receipt_number} - {str(e)}")
                
                # Progress every 5 receipts
                if i % 5 == 0:
                    elapsed = time.time() - start_time
                    rate = i / elapsed * 60  # per minute
                    logging.info(f"Progress: {i}/{total} | Success: {successful} | Failed: {failed} | Rate: {rate:.1f}/min")
                
                # Delay between receipts
                time.sleep(3)
            
            # Final summary
            elapsed_total = time.time() - start_time
            success_rate = (successful / total) * 100 if total > 0 else 0
            
            logging.info("=" * 50)
            logging.info("AUTOMATION COMPLETE")
            logging.info(f"Total Receipts: {total}")
            logging.info(f"Successful: {successful}")
            logging.info(f"Failed: {failed}")
            logging.info(f"Success Rate: {success_rate:.1f}%")
            logging.info(f"Total Time: {elapsed_total/60:.1f} minutes")
            logging.info(f"Processing Rate: {total/(elapsed_total/60):.1f} receipts/minute")
            logging.info(f"Processed by: deepanshudagdi")
            logging.info(f"Completed at: {time.strftime('%Y-%m-%d %H:%M:%S UTC')}")
            logging.info("=" * 50)
            
        except Exception as e:
            logging.error(f"FATAL ERROR: {str(e)}")
            
        finally:
            logging.info("Closing automation...")
            self.driver.quit()

if __name__ == "__main__":
    processor = EmitraCleanAutomation()
    processor.run_automation()