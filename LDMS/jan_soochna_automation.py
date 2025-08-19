import requests
import json
import time
import logging
from typing import List, Set
from dataclasses import dataclass
import re
import random
from google.oauth2 import service_account
from googleapiclient.discovery import build
import os

# Fixed logging setup for Windows
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('jan_soochna_automation.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

@dataclass
class BeneficiaryData:
    aadhaar_number: str
    name: str = ""
    father_name: str = ""
    address: str = ""
    gender: str = ""
    authority: str = ""
    renewal_date: str = ""
    registration_fees: str = ""
    application_status: str = ""
    application_number: str = ""
    card_issued_date: str = ""
    benefit_name: str = ""
    amount: str = ""
    bank_name: str = ""
    debit_date: str = ""
    apply_date: str = ""
    error_message: str = ""
    fetch_status: str = "Pending"

class JanSoochnaPortalClient:
    def __init__(self):
        pass

    def fetch_beneficiary_data(self, aadhaar_number: str) -> BeneficiaryData:
        """Fetch data from Jan Soochna portal"""
        
        logger.info(f"PROCESSING Aadhaar: {aadhaar_number}")
        beneficiary = BeneficiaryData(aadhaar_number=aadhaar_number)
        
        # Create fresh session
        session = requests.Session()
        session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'application/json, text/javascript, */*; q=0.01',
            'Accept-Language': 'en-US,en;q=0.9,hi;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'DNT': '1',
            'Connection': 'keep-alive'
        })
        
        BASE_URL = "https://jansoochna.rajasthan.gov.in"
        FORM_URL = "/Services/DynamicControlsDataSet"
        
        try:
            # Step 1: Get form page
            response = session.get(f"{BASE_URL}{FORM_URL}")
            
            if response.status_code != 200:
                logger.error(f"ERROR: Form page failed for {aadhaar_number}: {response.status_code}")
                beneficiary.error_message = f"Form page error: {response.status_code}"
                beneficiary.fetch_status = "Failed"
                self._fill_na_fields(beneficiary)
                return beneficiary
            
            # Step 2: Extract CSRF token
            csrf_match = re.search(r'name="__RequestVerificationToken"[^>]*value="([^"]+)"', response.text)
            if not csrf_match:
                logger.error(f"ERROR: No CSRF token for {aadhaar_number}")
                beneficiary.error_message = "No CSRF token found"
                beneficiary.fetch_status = "Failed"
                self._fill_na_fields(beneficiary)
                return beneficiary
            
            csrf_token = csrf_match.group(1)
            
            # Step 3: Prepare complete form data
            form_data = {
                '__RequestVerificationToken': csrf_token,
                'serviceID': '5s6nLtXarUM=',
                'machineId': '',
                'ipAddress': '',
                '_ListDynamicControlParent[0].EncryptedID': 'AScGcOfnRUA=',
                '_ListDynamicControlParent[0].DateFormat': '',
                '_ListDynamicControlParent[0].Control_Type': 'RADIO',
                '_ListDynamicControlParent[0].EncryptedValue': '',
                '_ListDynamicControlParent[0].Submit_Sequence': '0',
                '_ListDynamicControlParent[0].English_Control_Name': 'प्रकार चुनें',
                'रजिस्ट्रेशन नंबर': '',
                'आधार नंबर': '',
                'प्रकार_चुनें': 'False',
                'जन-आधार नंबर': '',
                '_ListDynamicControlParent[1].EncryptedID': 'UEecLGeEG1g=',
                '_ListDynamicControlParent[1].DateFormat': '',
                '_ListDynamicControlParent[1].Control_Type': 'TEXTBOX',
                '_ListDynamicControlParent[1].EncryptedValue': '',
                '_ListDynamicControlParent[1].Submit_Sequence': '4',
                '_ListDynamicControlParent[1].English_Control_Name': 'आई डी नंबर दर्ज़ करें',
                '_ListDynamicControlParent[1].ControlValue': aadhaar_number,
                '_ListDynamicControlParent[2].EncryptedID': 'yXsOYFWCGD0=',
                '_ListDynamicControlParent[2].DateFormat': '',
                '_ListDynamicControlParent[2].Control_Type': 'BUTTON',
                '_ListDynamicControlParent[2].EncryptedValue': '',
                '_ListDynamicControlParent[2].Submit_Sequence': '7',
                '_ListDynamicControlParent[2].English_Control_Name': 'खोजें',
                'seletedValue': 'yXsOYFWCGD0=',
                'value': 'c2dXlstdFew=',
                'selectedClass': '08DTNHkX+SU=',
                'RequiredValue': 'false',
                'SelectedValue': f'आई डी नंबर दर्ज़ करें:{aadhaar_number}'
            }
            
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'Origin': BASE_URL,
                'Referer': f"{BASE_URL}{FORM_URL}",
                'X-Requested-With': 'XMLHttpRequest'
            }
            
            # Step 4: Submit form
            response = session.post(f"{BASE_URL}{FORM_URL}", data=form_data, headers=headers, timeout=30)
            
            if response.status_code == 200:
                try:
                    # Parse JSON response
                    first_parse = json.loads(response.text)
                    if isinstance(first_parse, str):
                        data = json.loads(first_parse)
                    else:
                        data = first_parse
                    
                    # Extract Labour data
                    if isinstance(data, dict) and 'Labour' in data and data['Labour']:
                        labour_list = data['Labour']
                        if isinstance(labour_list, list) and len(labour_list) > 0:
                            labour_data = labour_list[0]
                            if isinstance(labour_data, dict):
                                # Extract all fields
                                beneficiary.name = str(labour_data.get('व्यक्ति / लाभार्थी का नाम / Beneficiary Name', '')).strip()
                                beneficiary.father_name = str(labour_data.get('व्यक्ति / लाभार्थी के पिता का नाम / Beneficiary Father Name ', '')).strip()
                                beneficiary.address = str(labour_data.get('व्यक्ति / लाभार्थी का पता / Address', '')).strip()
                                beneficiary.gender = str(labour_data.get('लिंग / Gender', '')).strip()
                                beneficiary.authority = str(labour_data.get('संबंधित प्राधिकरण / Concerned Union/Authority/Person', '')).strip()
                                beneficiary.renewal_date = str(labour_data.get('वैधता दिनांक / Renewal Due Date', '')).strip()
                                beneficiary.registration_fees = str(labour_data.get('आवेदन का शुल्क / Registration Fees ', '')).strip()
                                beneficiary.application_status = str(labour_data.get('आवेदन की स्थिति / Application Status ', '')).strip()
                                beneficiary.application_number = str(labour_data.get('आवेदन क्रमांक / Application Number ', '')).strip()
                                beneficiary.card_issued_date = str(labour_data.get('कार्ड जारी करने की दिनांक / Card Issued Date ', '')).strip()
                                
                                if beneficiary.name:
                                    beneficiary.fetch_status = "Success"
                                    logger.info(f"SUCCESS: {beneficiary.name} (Aadhaar: {aadhaar_number})")
                                    return beneficiary
                    
                    # No data found
                    beneficiary.error_message = "No beneficiary data found"
                    beneficiary.fetch_status = "Failed"
                    self._fill_na_fields(beneficiary)
                    return beneficiary
                    
                except json.JSONDecodeError as e:
                    beneficiary.error_message = f"JSON parsing failed: {str(e)}"
                    beneficiary.fetch_status = "Failed"
                    self._fill_na_fields(beneficiary)
                    return beneficiary
            else:
                beneficiary.error_message = f"HTTP Error: {response.status_code}"
                beneficiary.fetch_status = "Failed"
                self._fill_na_fields(beneficiary)
                return beneficiary
                
        except Exception as e:
            beneficiary.error_message = f"Error: {str(e)}"
            beneficiary.fetch_status = "Failed"
            self._fill_na_fields(beneficiary)
            return beneficiary

    def _fill_na_fields(self, beneficiary: BeneficiaryData):
        """Fill empty fields with N/A"""
        for field in beneficiary.__dataclass_fields__:
            if field not in ["aadhaar_number", "fetch_status", "error_message"]:
                current_value = getattr(beneficiary, field, "")
                if not current_value or current_value.strip() == "":
                    setattr(beneficiary, field, "N/A")

class GoogleSheetsManager:
    def __init__(self, credentials_file: str, spreadsheet_id: str):
        self.credentials_file = credentials_file
        self.spreadsheet_id = spreadsheet_id
        self.service = None
        self._initialize_service()

    def _initialize_service(self):
        try:
            credentials = service_account.Credentials.from_service_account_file(
                self.credentials_file,
                scopes=['https://www.googleapis.com/auth/spreadsheets']
            )
            self.service = build('sheets', 'v4', credentials=credentials)
            logger.info("Google Sheets service initialized")
        except Exception as e:
            logger.error(f"Failed to initialize Google Sheets: {e}")
            raise

    def ensure_header_exists(self, sheet_name: str):
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"{sheet_name}!A1:R1"
            ).execute()
            
            if not result.get('values'):
                header = [
                    "Aadhaar Number", "Name", "Father Name", "Address", "Gender", "Authority",
                    "Renewal Date", "Registration Fees", "Application Status", "Application Number", 
                    "Card Issued Date", "Benefit Name", "Amount", "Bank Name", "Debit Date",
                    "Apply Date", "Fetch Status", "Error Message"
                ]
                self.service.spreadsheets().values().update(
                    spreadsheetId=self.spreadsheet_id,
                    range=f"{sheet_name}!A1",
                    valueInputOption='RAW',
                    body={'values': [header]}
                ).execute()
                logger.info(f"Created header in {sheet_name}")
        except Exception as e:
            logger.error(f"Error ensuring header: {e}")

    def create_sheet_if_not_exists(self, sheet_name: str):
        try:
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
            existing_sheets = [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
            
            if sheet_name not in existing_sheets:
                request = {'addSheet': {'properties': {'title': sheet_name}}}
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id,
                    body={'requests': [request]}
                ).execute()
                logger.info(f"Created sheet: {sheet_name}")
                self.ensure_header_exists(sheet_name)
        except Exception as e:
            logger.error(f"Error creating sheet: {e}")

    def read_aadhaar_numbers(self, sheet_name: str = "Sheet1", column: str = "A") -> List[str]:
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"{sheet_name}!{column}:{column}"
            ).execute()
            
            values = result.get('values', [])
            aadhaar_numbers = []
            
            for row in values:
                if row and len(row) > 0:
                    aadhaar = str(row[0]).strip()
                    if len(aadhaar) == 12 and aadhaar.isdigit():
                        aadhaar_numbers.append(aadhaar)
            
            logger.info(f"Read {len(aadhaar_numbers)} valid Aadhaar numbers from {sheet_name}")
            return aadhaar_numbers
        except Exception as e:
            logger.error(f"Error reading Aadhaar numbers: {e}")
            return []

    def read_existing_results(self, sheet_name: str = "Results") -> Set[str]:
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"{sheet_name}!A2:A"
            ).execute()
            
            values = result.get('values', [])
            existing = set()
            
            for row in values:
                if row and len(row) > 0:
                    aadhaar = str(row[0]).strip()
                    if len(aadhaar) == 12 and aadhaar.isdigit():
                        existing.add(aadhaar)
            
            return existing
        except Exception as e:
            logger.error(f"Error reading existing results: {e}")
            return set()

    def write_result(self, beneficiary: BeneficiaryData, sheet_name: str = "Results"):
        try:
            self.create_sheet_if_not_exists(sheet_name)
            self.ensure_header_exists(sheet_name)
            
            # Prepare row data
            row = [
                beneficiary.aadhaar_number,
                beneficiary.name or "N/A",
                beneficiary.father_name or "N/A", 
                beneficiary.address or "N/A",
                beneficiary.gender or "N/A",
                beneficiary.authority or "N/A",
                beneficiary.renewal_date or "N/A",
                beneficiary.registration_fees or "N/A",
                beneficiary.application_status or "N/A",
                beneficiary.application_number or "N/A",
                beneficiary.card_issued_date or "N/A",
                beneficiary.benefit_name or "N/A",
                beneficiary.amount or "N/A",
                beneficiary.bank_name or "N/A",
                beneficiary.debit_date or "N/A",
                beneficiary.apply_date or "N/A",
                beneficiary.fetch_status,
                beneficiary.error_message or ""
            ]
            
            # Append to sheet
            self.service.spreadsheets().values().append(
                spreadsheetId=self.spreadsheet_id,
                range=f"{sheet_name}!A2",
                valueInputOption='RAW',
                insertDataOption='INSERT_ROWS',
                body={'values': [row]}
            ).execute()
            
            logger.info(f"WRITTEN: Result for {beneficiary.aadhaar_number}")
            return True
            
        except Exception as e:
            logger.error(f"Error writing result: {e}")
            return False

    def clear_results_sheet(self, sheet_name: str = "Results"):
        """Clear the results sheet"""
        try:
            self.service.spreadsheets().values().clear(
                spreadsheetId=self.spreadsheet_id,
                range=f"{sheet_name}!A:Z",
                body={}
            ).execute()
            logger.info(f"Cleared {sheet_name} sheet")
            return True
        except Exception as e:
            logger.error(f"Error clearing sheet: {e}")
            return False

    def show_results(self, sheet_name: str = "Results"):
        """Display current results"""
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=f"{sheet_name}!A:R"
            ).execute()
            
            values = result.get('values', [])
            
            if not values:
                print("No data found in Results sheet")
                return
            
            print(f"\nFound {len(values)} rows in Results sheet")
            print("=" * 80)
            
            # Show summary
            success_count = 0
            failed_count = 0
            
            for i, row in enumerate(values):
                if i == 0:  # Header
                    continue
                
                if len(row) >= 17:
                    aadhaar = row[0]
                    name = row[1]
                    status = row[16]
                    error_msg = row[17] if len(row) > 17 else ""
                    
                    print(f"Aadhaar: {aadhaar} | Name: {name} | Status: {status}")
                    if status == "Success":
                        success_count += 1
                    elif status == "Failed":
                        failed_count += 1
                        if error_msg:
                            print(f"  Error: {error_msg}")
            
            print("=" * 80)
            print(f"SUMMARY: {success_count} successful, {failed_count} failed, {len(values)-1} total")
            
        except Exception as e:
            print(f"Error reading results: {e}")

class JanSoochnaAutomation:
    def __init__(self, credentials_file: str, spreadsheet_id: str, delay_seconds: int = 6):
        self.portal_client = JanSoochnaPortalClient()
        self.sheets_manager = GoogleSheetsManager(credentials_file, spreadsheet_id)
        self.delay_seconds = delay_seconds

    def run_automation(self, input_sheet: str = "Sheet1", input_column: str = "A", output_sheet: str = "Results"):
        logger.info("STARTING Jan Soochna automation")
        
        try:
            # Read Aadhaar numbers
            aadhaar_numbers = self.sheets_manager.read_aadhaar_numbers(input_sheet, input_column)
            if not aadhaar_numbers:
                logger.warning("No Aadhaar numbers found")
                return []
            
            # Check existing results
            existing = self.sheets_manager.read_existing_results(output_sheet)
            to_process = [a for a in aadhaar_numbers if a not in existing]
            
            logger.info(f"SUMMARY: Total={len(aadhaar_numbers)}, Already done={len(existing)}, To process={len(to_process)}")
            
            if not to_process:
                logger.info("ALL Aadhaar numbers already processed")
                return []
            
            # Process each Aadhaar
            results = []
            success_count = 0
            failed_count = 0
            
            for i, aadhaar in enumerate(to_process, 1):
                logger.info(f"PROCESSING {i}/{len(to_process)}: {aadhaar}")
                
                # Fetch data
                result = self.portal_client.fetch_beneficiary_data(aadhaar)
                results.append(result)
                
                # Write immediately
                success = self.sheets_manager.write_result(result, output_sheet)
                
                if success and result.fetch_status == "Success":
                    success_count += 1
                    logger.info(f"SUCCESS: {result.name}")
                elif success and result.fetch_status == "Failed":
                    failed_count += 1
                    logger.info(f"FAILED: {result.error_message}")
                
                # Delay before next
                if i < len(to_process):
                    delay = random.uniform(self.delay_seconds, self.delay_seconds + 3)
                    logger.info(f"WAITING {delay:.1f} seconds...")
                    time.sleep(delay)
            
            # Final summary
            logger.info(f"COMPLETED: {success_count} successful, {failed_count} failed")
            return results
            
        except Exception as e:
            logger.error(f"Automation error: {e}")
            return []

def show_menu():
    """Display main menu"""
    print("\n" + "="*60)
    print("         JAN SOOCHNA AUTOMATION SYSTEM")
    print("="*60)
    print("1. Run Automation (Process Aadhaar numbers)")
    print("2. View Results")
    print("3. Clear Results (to reprocess all)")
    print("4. Test Single Aadhaar")
    print("5. Exit")
    print("="*60)

def test_single_aadhaar():
    """Test single Aadhaar"""
    aadhaar = input("Enter Aadhaar number to test: ").strip()
    if len(aadhaar) != 12 or not aadhaar.isdigit():
        print("ERROR: Please enter a valid 12-digit Aadhaar number")
        return
    
    client = JanSoochnaPortalClient()
    result = client.fetch_beneficiary_data(aadhaar)
    
    print(f"\nTEST RESULT:")
    print(f"Aadhaar: {result.aadhaar_number}")
    print(f"Status: {result.fetch_status}")
    if result.fetch_status == "Success":
        print(f"Name: {result.name}")
        print(f"Father Name: {result.father_name}")
        print(f"Address: {result.address}")
    else:
        print(f"Error: {result.error_message}")

def main():
    """Main application"""
    # Configuration
    CREDENTIALS_FILE = "google_sheets_credentials.json"
    SPREADSHEET_ID = "1y1fnk7dGjZAg3gasp7njQHWSwAeGE-0NXNFZCrdtK_E"
    
    # Check if credentials file exists
    if not os.path.exists(CREDENTIALS_FILE):
        print(f"ERROR: {CREDENTIALS_FILE} not found!")
        print("Please ensure your Google Sheets credentials file is in the current directory.")
        return
    
    try:
        automation = JanSoochnaAutomation(CREDENTIALS_FILE, SPREADSHEET_ID, delay_seconds=6)
        
        while True:
            show_menu()
            choice = input("Enter your choice (1-5): ").strip()
            
            if choice == "1":
                print("\nStarting automation...")
                results = automation.run_automation("Sheet1", "A", "Results")
                if results:
                    success_count = len([r for r in results if r.fetch_status == "Success"])
                    print(f"\nAutomation completed: {success_count}/{len(results)} successful")
                else:
                    print("No new Aadhaar numbers to process")
                input("\nPress Enter to continue...")
                
            elif choice == "2":
                print("\nCurrent Results:")
                automation.sheets_manager.show_results("Results")
                input("\nPress Enter to continue...")
                
            elif choice == "3":
                confirm = input("Are you sure you want to clear all results? (yes/no): ")
                if confirm.lower() == 'yes':
                    if automation.sheets_manager.clear_results_sheet("Results"):
                        print("Results cleared successfully!")
                    else:
                        print("Failed to clear results")
                else:
                    print("Operation cancelled")
                input("\nPress Enter to continue...")
                
            elif choice == "4":
                test_single_aadhaar()
                input("\nPress Enter to continue...")
                
            elif choice == "5":
                print("Goodbye!")
                break
                
            else:
                print("Invalid choice. Please try again.")
                
    except Exception as e:
        logger.error(f"Application error: {e}")
        print(f"ERROR: {e}")

if __name__ == "__main__":
    main()