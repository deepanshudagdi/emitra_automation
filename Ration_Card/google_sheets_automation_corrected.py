import gspread
import time
from datetime import datetime
from portal_scraper import RajasthanFoodPortalScraper
import logging
import re

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class GoogleSheetsRationCardAutomation:
    def __init__(self, credentials_file='credentials.json'):
        self.credentials_file = credentials_file
        self.gc = None
        self.sheet = None
        self.scraper = RajasthanFoodPortalScraper(headless=True)
        
    def authenticate(self):
        """Authenticate with Google Sheets API"""
        try:
            self.gc = gspread.service_account(filename=self.credentials_file)
            logger.info("✅ Successfully authenticated with Google Sheets API")
            return True
        except Exception as e:
            logger.error(f"❌ Authentication failed: {str(e)}")
            return False
    
    def open_sheet(self, sheet_url_or_key, worksheet_name=None):
        """Open Google Sheet by URL or key"""
        try:
            if 'docs.google.com/spreadsheets' in sheet_url_or_key:
                match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', sheet_url_or_key)
                if match:
                    sheet_id = match.group(1)
                else:
                    raise ValueError("Could not extract sheet ID from URL")
            else:
                sheet_id = sheet_url_or_key
            
            spreadsheet = self.gc.open_by_key(sheet_id)
            
            if worksheet_name:
                self.sheet = spreadsheet.worksheet(worksheet_name)
            else:
                self.sheet = spreadsheet.sheet1
            
            logger.info(f"✅ Successfully opened sheet: {self.sheet.title}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Failed to open sheet: {str(e)}")
            return False
    
    def setup_headers(self):
        """Setup column headers - CORRECTED VERSION"""
        try:
            headers = self.sheet.row_values(1)
            
            # Updated headers - removed timestamp and search status columns
            expected_headers = [
                'Ration Card Number',  # Column A (input)
                'Office Name',         # Column B
                'Form Number',         # Column C
                'Token Number',        # Column D
                'User ID',             # Column E
                'Status'               # Column F
            ]
            
            if len(headers) < len(expected_headers):
                self.sheet.update('A1:F1', [expected_headers])
                logger.info("✅ Headers setup completed (6 columns)")
                time.sleep(1)
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Failed to setup headers: {str(e)}")
            return False
    
    def get_ration_card_numbers(self, start_row=2):
        """Get ration card numbers from column A"""
        try:
            all_values = self.sheet.get_all_values()
            
            if len(all_values) < start_row:
                logger.info("No data rows found")
                return []
            
            ration_numbers = []
            for i in range(start_row - 1, len(all_values)):
                row_data = all_values[i]
                if row_data and len(row_data) > 0 and row_data[0]:
                    ration_number = str(row_data[0]).strip()
                    if ration_number and ration_number.lower() not in ['', 'nan', 'none']:
                        ration_numbers.append({
                            'row': i + 1,
                            'number': ration_number
                        })
            
            logger.info(f"📋 Found {len(ration_numbers)} ration card numbers to process")
            return ration_numbers
            
        except Exception as e:
            logger.error(f"❌ Failed to get ration card numbers: {str(e)}")
            return []
    
    def parse_search_result(self, result):
        """Parse scraper result with CORRECTED User ID extraction"""
        parsed = {
            'office_name': '',
            'form_number': '',
            'token_number': '',
            'user_id': '',
            'status': ''
        }
        
        if 'error' in result:
            return parsed  # Return empty values for errors
        
        try:
            # Look for the main data in all table content
            data_found = False
            
            # Check all content for ration card data
            for key, content in result.items():
                if not content or not isinstance(content, str):
                    continue
                    
                # Skip contact information
                if 'Contact No' in content or 'Email' in content or 'Address' in content:
                    continue
                    
                lines = content.split('\n')
                for line in lines:
                    # Look for the main data line with office name and numbers
                    if ('प्राधिकृत अधिकारी' in line or 'Officer' in line) and any(c.isdigit() for c in line):
                        data_found = True
                        
                        # Clean and parse the line
                        clean_line = line.replace('*', '').strip()
                        parts = clean_line.split()
                        
                        logger.info(f"🔍 Parsing line: {clean_line}")
                        
                        # IMPROVED EXTRACTION LOGIC
                        office_parts = []
                        numbers = []
                        user_id = ''
                        status_parts = []
                        
                        i = 0
                        while i < len(parts):
                            part = parts[i]
                            
                            # Extract numbers (Form Number and Token Number)
                            if part.isdigit() and len(part) >= 8:
                                numbers.append(part)
                                logger.info(f"📊 Found number: {part}")
                            
                            # Extract User ID (format: Letter + numbers, e.g., K119269051)
                            elif (len(part) > 5 and part[0].isalpha() and 
                                  part[1:].isdigit() and not 'Printed' in part):
                                user_id = part
                                logger.info(f"👤 Found User ID: {part}")
                            
                            # Extract Status (Ration Card Printed(...))
                            elif part == 'Ration' and i + 2 < len(parts):
                                if parts[i+1] == 'Card' and 'Printed' in parts[i+2]:
                                    # Look for the complete status including parentheses
                                    status_start = i
                                    j = i
                                    while j < len(parts) and not parts[j].endswith(')'):
                                        j += 1
                                    if j < len(parts):
                                        j += 1  # Include the closing parenthesis
                                    status_parts = parts[status_start:j]
                                    break
                            
                            # Extract Office Name (text that's not numbers, user ID, or status)
                            elif (not part.isdigit() and 
                                  not (len(part) > 5 and part[0].isalpha() and part[1:].isdigit()) and
                                  part not in ['Ration', 'Card', 'Printed', 'Form', 'Token', 'User', 'Status']):
                                office_parts.append(part)
                            
                            i += 1
                        
                        # Assign parsed values
                        if office_parts:
                            parsed['office_name'] = ' '.join(office_parts)
                            logger.info(f"🏢 Office: {parsed['office_name']}")
                        
                        if len(numbers) >= 2:
                            parsed['form_number'] = numbers[0]
                            parsed['token_number'] = numbers[1]
                            logger.info(f"📋 Form: {parsed['form_number']}")
                            logger.info(f"🎫 Token: {parsed['token_number']}")
                        elif len(numbers) == 1:
                            parsed['form_number'] = numbers[0]
                            logger.info(f"📋 Form: {parsed['form_number']}")
                        
                        if user_id:
                            parsed['user_id'] = user_id
                            logger.info(f"👤 User ID: {parsed['user_id']}")
                        
                        if status_parts:
                            parsed['status'] = ' '.join(status_parts)
                            logger.info(f"📊 Status: {parsed['status']}")
                        else:
                            # Fallback: look for "Printed" pattern in the original line
                            if 'Printed' in line:
                                status_match = re.search(r'Ration Card Printed\([^)]+\)', line)
                                if status_match:
                                    parsed['status'] = status_match.group()
                                    logger.info(f"📊 Status (regex): {parsed['status']}")
                        
                        break
                
                if data_found:
                    break
            
        except Exception as e:
            logger.error(f"Parse error: {str(e)}")
        
        return parsed
    
    def update_row_data(self, row_number, parsed_data):
        """Update a single row with parsed data - CORRECTED to 5 columns only"""
        try:
            # Prepare data for columns B to F (5 columns total)
            row_data = [
                parsed_data['office_name'],
                parsed_data['form_number'], 
                parsed_data['token_number'],
                parsed_data['user_id'],
                parsed_data['status']
            ]
            
            # Update columns B to F only
            range_name = f'B{row_number}:F{row_number}'
            self.sheet.update(values=[row_data], range_name=range_name)
            time.sleep(1)
            
            logger.info(f"✅ Updated row {row_number} with search results")
            return True
            
        except Exception as e:
            logger.error(f"❌ Failed to update row {row_number}: {str(e)}")
            return False
    
    def process_all_ration_cards(self, start_row=2, delay_seconds=5):
        """Process all ration card numbers"""
        try:
            self.scraper.start_driver()
            ration_numbers = self.get_ration_card_numbers(start_row)
            
            if not ration_numbers:
                logger.info("No ration card numbers found to process")
                return True
            
            processed_count = 0
            success_count = 0
            
            print(f"\n🚀 Starting to process {len(ration_numbers)} ration card numbers...")
            print("=" * 60)
            
            for item in ration_numbers:
                row_num = item['row']
                ration_number = item['number']
                
                print(f"\n📋 Processing {processed_count + 1}/{len(ration_numbers)}")
                print(f"🔢 Ration Card: {ration_number} (Row {row_num})")
                
                # Search with retry logic
                max_retries = 2
                search_result = None
                
                for attempt in range(max_retries):
                    try:
                        search_result = self.scraper.search_ration_card(ration_number)
                        if 'error' not in search_result:
                            break
                        elif attempt < max_retries - 1:
                            print(f"⚠️  Attempt {attempt + 1} failed, retrying...")
                            time.sleep(3)
                    except Exception as e:
                        if attempt < max_retries - 1:
                            print(f"⚠️  Attempt {attempt + 1} failed: {str(e)}, retrying...")
                            time.sleep(3)
                        else:
                            search_result = {"error": f"All attempts failed: {str(e)}"}
                
                if not search_result:
                    search_result = {"error": "No response from portal"}
                
                # Parse results
                parsed_data = self.parse_search_result(search_result)
                
                # Update sheet
                if self.update_row_data(row_num, parsed_data):
                    success_count += 1
                    
                    # Show results
                    if any([parsed_data['office_name'], parsed_data['form_number'], 
                           parsed_data['token_number'], parsed_data['user_id'], parsed_data['status']]):
                        print(f"✅ Success! Found:")
                        if parsed_data['office_name']:
                            print(f"   🏢 Office: {parsed_data['office_name']}")
                        if parsed_data['form_number']:
                            print(f"   📋 Form Number: {parsed_data['form_number']}")
                        if parsed_data['token_number']:
                            print(f"   🎫 Token Number: {parsed_data['token_number']}")
                        if parsed_data['user_id']:
                            print(f"   👤 User ID: {parsed_data['user_id']}")
                        if parsed_data['status']:
                            print(f"   📊 Status: {parsed_data['status']}")
                    else:
                        print(f"⚠️  No data found for this ration card")
                else:
                    print(f"❌ Failed to update sheet for row {row_num}")
                
                processed_count += 1
                
                if processed_count < len(ration_numbers):
                    print(f"⏳ Waiting {delay_seconds} seconds...")
                    time.sleep(delay_seconds)
            
            print(f"\n🎉 Processing completed!")
            print(f"📊 Summary:")
            print(f"   - Total processed: {processed_count}")
            print(f"   - Successfully updated: {success_count}")
            print(f"   - Failed: {processed_count - success_count}")
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Error during processing: {str(e)}")
            return False
        finally:
            self.scraper.close()
    
    def run_automation(self, sheet_url, worksheet_name=None, start_row=2):
        """Complete automation workflow"""
        print("🤖 Google Sheets Ration Card Automation - CORRECTED VERSION")
        print("=" * 55)
        
        print("🔐 Step 1: Authenticating with Google Sheets...")
        if not self.authenticate():
            return False
        
        print("📊 Step 2: Opening Google Sheet...")
        if not self.open_sheet(sheet_url, worksheet_name):
            return False
        
        print("📝 Step 3: Setting up headers (6 columns only)...")
        if not self.setup_headers():
            return False
        
        print("🚀 Step 4: Processing ration card numbers...")
        if not self.process_all_ration_cards(start_row):
            return False
        
        print("\n✅ Automation completed successfully! 🎉")
        return True

def run_sheets_automation(sheet_url, worksheet_name=None):
    """Main function to run the corrected automation"""
    automation = GoogleSheetsRationCardAutomation()
    return automation.run_automation(sheet_url, worksheet_name)

if __name__ == "__main__":
    # Run with your sheet details
    sheet_url = "16l3w3hcGAVq2MoB_bP1hvfDHKYxaV1N1SnS6K4ywx0M"
    worksheet_name = "Ration Card"
    run_sheets_automation(sheet_url, worksheet_name)