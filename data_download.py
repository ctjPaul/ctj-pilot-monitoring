"""
Uplink Data Downloader Module
Handles automated login and data download from Uplink web portal
INTEGRATED WITH WORKING UPLINK EXPORT SCRIPT
"""

import os
import time
import shutil
import glob
from datetime import datetime
from pathlib import Path
import pandas as pd
import xml.etree.ElementTree as ET
import streamlit as st

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException



class UplinkDownloader:
    """Handles automated data download from Uplink portal"""
    
    def __init__(self, download_dir=None, headless=True):
        """
        Initialize the Uplink downloader
        
        Args:
            download_dir (str): Directory for downloads. Uses default if None
            headless (bool): Run browser in headless mode (default False for debugging)
        """
        self.download_dir = download_dir or r"C:\Users\Paul\OneDrive - CTJ Energy\Python Scripts\Monthly_Reports"
        Path(self.download_dir).mkdir(parents=True, exist_ok=True)
        self.headless = headless
        self.driver = None
        
        # Uplink credentials
        self.uplink_url = "https://uplinkdealers.com/dealers/login"
        # NEW (SECURE):
	
	try:
    	self.username = st.secrets["UPLINK_USERNAME"]
    	self.password = st.secrets["UPLINK_PASSWORD"]
	except:
    	# Fallback for local testing
   	 self.username = os.getenv("UPLINK_USERNAME", "")
    	self.password = os.getenv("UPLINK_PASSWORD", "")
    
    def convert_xml_to_excel(self, xml_file_path, excel_file_path):
        """Convert Microsoft Office XML Spreadsheet format to modern Excel"""
        print(f"Converting {xml_file_path} to modern Excel format...")
        
        try:
            with open(xml_file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            print(f"📄 File content preview (first 200 chars):")
            print(content[:200])
            
            # Check if this is Microsoft Office XML Spreadsheet format
            if 'xmlns="urn:schemas-microsoft-com:office:spreadsheet"' in content:
                print("📋 Detected Microsoft Office XML Spreadsheet format")
                
                try:
                    root = ET.fromstring(content)
                    
                    # Define the namespace
                    ns = {
                        'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
                        'o': 'urn:schemas-microsoft-com:office:office',
                        'x': 'urn:schemas-microsoft-com:office:excel',
                        'html': 'http://www.w3.org/TR/REC-html40'
                    }
                    
                    # Find the worksheet and table
                    worksheet = root.find('.//ss:Worksheet', ns)
                    if worksheet is None:
                        print("❌ No worksheet found in XML")
                        return False
                    
                    table = worksheet.find('.//ss:Table', ns)
                    if table is None:
                        print("❌ No table found in worksheet")
                        return False
                    
                    # Extract rows
                    rows = table.findall('.//ss:Row', ns)
                    if not rows:
                        print("❌ No rows found in table")
                        return False
                    
                    print(f"📊 Found {len(rows)} rows in the spreadsheet")
                    
                    # Extract data
                    data_rows = []
                    headers = []
                    
                    for row_index, row in enumerate(rows):
                        cells = row.findall('.//ss:Cell', ns)
                        row_data = []
                        
                        for cell in cells:
                            data_elem = cell.find('.//ss:Data', ns)
                            if data_elem is not None and data_elem.text:
                                row_data.append(data_elem.text.strip())
                            else:
                                row_data.append('')
                        
                        if row_index == 0:
                            # First row is headers
                            headers = row_data
                            print(f"📋 Column headers: {headers}")
                        else:
                            # Data rows
                            if len(row_data) == len(headers):
                                data_dict = {headers[i]: row_data[i] for i in range(len(headers))}
                                data_rows.append(data_dict)
                    
                    if data_rows:
                        # Create DataFrame and save as Excel
                        df = pd.DataFrame(data_rows)
                        df.to_excel(excel_file_path, index=False, engine='openpyxl')
                        print(f"✅ Successfully converted Microsoft Office XML with {len(data_rows)} data rows")
                        print(f"📊 Columns: {list(df.columns)}")
                        return True
                    else:
                        print("❌ No data rows found")
                        return False
                        
                except ET.ParseError as e:
                    print(f"❌ XML parsing failed: {e}")
                    return False
            
            # Fallback: Try other methods for different XML formats
            elif content.strip().startswith('<?xml'):
                print("📋 Detected generic XML format, attempting standard parsing...")
                
                try:
                    root = ET.fromstring(content)
                    
                    # Try to extract tabular data from generic XML
                    data_rows = []
                    
                    # Look for repeating elements that might represent rows
                    for elem in root.iter():
                        if len(list(elem)) > 0:  # Has child elements
                            row_data = {}
                            for child in elem:
                                if child.text and child.text.strip():
                                    row_data[child.tag] = child.text.strip()
                            
                            if len(row_data) > 1:  # Only add rows with multiple columns
                                data_rows.append(row_data)
                    
                    if data_rows:
                        df = pd.DataFrame(data_rows)
                        df.to_excel(excel_file_path, index=False, engine='openpyxl')
                        print(f"✅ Successfully converted generic XML with {len(data_rows)} rows")
                        return True
                        
                except Exception as e:
                    print(f"⚠️ Generic XML parsing failed: {e}")
            
            # Fallback: Try reading as delimited data
            else:
                print("📋 Attempting to parse as delimited data...")
                
                # Try different delimiters
                for delimiter in ['\t', ',', '|', ';']:
                    try:
                        df = pd.read_csv(xml_file_path, delimiter=delimiter, encoding='utf-8')
                        if len(df.columns) > 1:
                            df.to_excel(excel_file_path, index=False, engine='openpyxl')
                            print(f"✅ Successfully converted as {repr(delimiter)}-delimited data")
                            return True
                    except:
                        continue
            
            print("❌ All conversion methods failed")
            return False
            
        except Exception as e:
            print(f"❌ Conversion error: {e}")
            return False
    
    def setup_driver(self):
        """Configure and initialize the Chrome WebDriver"""
        
        chrome_options = Options()
        
        # Set download directory
        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_settings.popups": 0,
            "profile.default_content_setting_values.automatic_downloads": 1,
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_argument("--disable-download-notification")
        
        if self.headless:
            chrome_options.add_argument("--headless")
        
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        
        # Initialize driver
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.implicitly_wait(10)
        
        return self.driver
    
    def login(self):
        """Login to Uplink portal"""
        
        try:
            print("Step 1: Logging in...")
            self.driver.get(self.uplink_url)
            wait = WebDriverWait(self.driver, 10)
            
            username_field = wait.until(EC.presence_of_element_located((By.ID, "Username")))
            password_field = self.driver.find_element(By.ID, "Password")
            
            username_field.send_keys(self.username)
            password_field.send_keys(self.password)
            
            self.driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()
            time.sleep(5)
            print("✓ Login successful!")
            return True
            
        except Exception as e:
            print(f"✗ Login failed: {e}")
            return False
    
    def navigate_to_device(self, device_id, device_name):
        """Navigate to specific device page"""
        
        try:
            print("Step 2: Navigating to devices...")
            try:
                devices_link = self.driver.find_element(By.PARTIAL_LINK_TEXT, "Devices")
                devices_link.click()
            except:
                try:
                    devices_link = self.driver.find_element(By.PARTIAL_LINK_TEXT, "Units")
                    devices_link.click()
                except:
                    self.driver.get("https://uplinkdealers.com/dealers/devices")
            
            time.sleep(3)
            
            print(f"Step 3: Finding {device_name}...")
            
            # Find all elements containing device info
            search_terms = [device_name, device_id, device_name.split('-')[-1]]
            device_element = None
            
            for term in search_terms:
                elements = self.driver.find_elements(By.XPATH, f"//*[contains(text(), '{term}')]")
                if elements:
                    print(f"Found {len(elements)} elements containing '{term}'")
                    device_element = elements[0]
                    break
            
            if not device_element:
                print(f"✗ Could not find device {device_name}")
                return False
            
            print(f"Found device element: {device_element.text[:50]}...")
            
            # Try to open device dashboard
            dashboard_opened = False
            
            # Try Alternative 1: Look for a direct link
            print("\nLooking for direct dashboard link...")
            try:
                if device_element.tag_name == 'a':
                    print("Device element is a link, clicking it...")
                    device_element.click()
                    time.sleep(3)
                    dashboard_opened = True
                else:
                    parent = device_element.find_element(By.XPATH, "..")
                    links = parent.find_elements(By.TAG_NAME, "a")
                    if links:
                        print(f"Found {len(links)} links near device")
                        for link in links:
                            if "dashboard" in link.get_attribute("href").lower():
                                print("Found dashboard link!")
                                link.click()
                                dashboard_opened = True
                                break
            except Exception as e:
                print(f"Direct link approach didn't work: {e}")
            
            # Try Alternative 2: Right-click with context menu
            if not dashboard_opened:
                print("\nTrying right-click approach...")
                try:
                    actions = ActionChains(self.driver)
                    actions.context_click(device_element).perform()
                    time.sleep(1)
                    
                    menu_selectors = [
                        "//ul[contains(@class, 'context')]//li",
                        "//div[contains(@class, 'menu')]//a",
                        "//li[contains(text(), 'Dashboard')]",
                        "//a[contains(text(), 'Dashboard')]"
                    ]
                    
                    for selector in menu_selectors:
                        menu_items = self.driver.find_elements(By.XPATH, selector)
                        if menu_items:
                            for item in menu_items:
                                if "dashboard" in item.text.lower():
                                    print(f"Clicking on: {item.text}")
                                    item.click()
                                    dashboard_opened = True
                                    break
                        if dashboard_opened:
                            break
                except Exception as e:
                    print(f"Right-click approach error: {e}")
            
            # Wait for dashboard to open
            time.sleep(5)
            
            # Check if new window opened
            if len(self.driver.window_handles) > 1:
                self.driver.switch_to.window(self.driver.window_handles[-1])
                print("Switched to new window")
            
            print(f"Current URL: {self.driver.current_url}")
            return True
            
        except Exception as e:
            print(f"✗ Failed to navigate to device: {e}")
            return False
    
    def navigate_to_events(self):
        """Navigate to Events page"""
        
        try:
            print("\nStep 4: Looking for Events link...")
            time.sleep(3)  # Wait for page to fully load
            
            events_clicked = False
            events_selectors = [
                "//a[text()='Events']",
                "//a[contains(text(), 'Events')]",
                "//span[contains(text(), 'Events')]",
                "//*[contains(@onclick, 'Events')]",
            ]
            
            for selector in events_selectors:
                try:
                    events_elements = self.driver.find_elements(By.XPATH, selector)
                    for events_link in events_elements:
                        if events_link.is_displayed():
                            print(f"Found Events element: {events_link.tag_name}")
                            try:
                                events_link.click()
                                events_clicked = True
                                print("✓ Successfully clicked Events")
                                break
                            except:
                                try:
                                    self.driver.execute_script("arguments[0].click();", events_link)
                                    events_clicked = True
                                    print("✓ Successfully clicked Events with JavaScript")
                                    break
                                except:
                                    continue
                    if events_clicked:
                        break
                except:
                    continue
            
            time.sleep(4)
            return events_clicked
            
        except Exception as e:
            print(f"✗ Failed to navigate to Events: {e}")
            return False
    
    def set_custom_date_range(self, start_date, end_date):
        """Set custom date range in the Events page - USING ORIGINAL WORKING LOGIC"""
        
        try:
            print("\n" + "="*50)
            print("CRITICAL: Selecting Custom from Intervals dropdown...")
            print("="*50)
            
            # Find the Intervals dropdown - try ALL the original selectors
            intervals_dropdown = None
            dropdown_selectors = [
                "//select[contains(@class, 'x-form-field')]",
                "//input[@value='Last 24 Hour']",
                "//div[contains(text(), 'Last 24 Hour')]",
                "//div[contains(@class, 'x-combo')]//input",
                "//div[contains(@class, 'x-form-field') and contains(., 'Last 24 Hour')]",
                "//input[contains(@class, 'x-form-field')]",
                "//*[text()='Intervals:']/following-sibling::*//input",
                "//*[text()='Intervals:']/following::input[1]",
                "//input[contains(@value, 'Hour')]"
            ]
            
            print("Searching for Intervals dropdown...")
            for selector in dropdown_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    print(f"Selector '{selector}' found {len(elements)} elements")
                    for elem in elements:
                        try:
                            elem_value = elem.get_attribute('value') or elem.text or ""
                            print(f"  - Element value: '{elem_value}', visible: {elem.is_displayed()}")
                            if elem.is_displayed() and ('24' in elem_value or 'Hour' in elem_value):
                                intervals_dropdown = elem
                                print(f"✓ Found Intervals dropdown with value: {elem_value}")
                                break
                        except:
                            continue
                    if intervals_dropdown:
                        break
                except Exception as e:
                    print(f"  Error with selector: {e}")
                    continue
            
            if not intervals_dropdown:
                print("❌ ERROR: Could not find Intervals dropdown!")
                return False
            
            # Click to open dropdown
            print("\nClicking to open Intervals dropdown...")
            try:
                intervals_dropdown.click()
                print("✓ Clicked with regular click")
            except:
                try:
                    self.driver.execute_script("arguments[0].click();", intervals_dropdown)
                    print("✓ Clicked with JavaScript")
                except Exception as e:
                    print(f"❌ Could not click dropdown: {e}")
                    return False
            
            time.sleep(1)
            
            # Now find and click "Custom" option - try ALL original selectors
            print("\nLooking for Custom option...")
            custom_selectors = [
                "//div[text()='Custom']",
                "//li[text()='Custom']",
                "//div[contains(@class, 'x-combo-list-item') and text()='Custom']",
                "//option[text()='Custom']",
                "//*[contains(@class, 'x-combo-list')]//div[text()='Custom']"
            ]
            
            custom_clicked = False
            for selector in custom_selectors:
                try:
                    custom_elements = self.driver.find_elements(By.XPATH, selector)
                    print(f"Selector '{selector}' found {len(custom_elements)} Custom elements")
                    for custom_option in custom_elements:
                        try:
                            if custom_option.is_displayed():
                                print(f"Found visible Custom option, clicking it...")
                                custom_option.click()
                                custom_clicked = True
                                print("✓ Selected Custom interval!")
                                break
                        except:
                            continue
                    if custom_clicked:
                        break
                except:
                    continue
            
            if not custom_clicked:
                print("❌ Could not click Custom from dropdown, trying keyboard method...")
                try:
                    # Clear and type Custom
                    intervals_dropdown.send_keys(Keys.CONTROL + "a")
                    intervals_dropdown.send_keys("Custom")
                    time.sleep(0.5)
                    intervals_dropdown.send_keys(Keys.ENTER)
                    print("✓ Typed 'Custom' into dropdown")
                except Exception as e:
                    print(f"❌ Keyboard method also failed: {e}")
                    return False
            
            # Wait for Custom to be selected and date fields to become active
            print("\nWaiting for date fields to become active...")
            time.sleep(2)
            
            # Now set the date range - USING ORIGINAL SCRIPT'S EXACT LOGIC
            print("\n" + "="*50)
            print("Setting date range...")
            print("="*50)
            
            start_str = start_date.strftime("%m/%d/%Y")
            end_str = end_date.strftime("%m/%d/%Y")
            print(f"Date range: {start_str} to {end_str}")
            
            # Find the From and To date fields - try ALL original selectors
            from_field = None
            to_field = None
            
            # Find From field
            from_selectors = [
                "//span[text()='From:']/following::input[1]",
                "//label[contains(text(), 'From')]/following::input[1]",
                "//*[text()='From:']/following-sibling::*//input",
                "//input[contains(@id, 'from')]",
                "//input[contains(@name, 'from')]",
                "//td[text()='From:']/following-sibling::td//input"
            ]
            
            print("\nSearching for FROM field...")
            for selector in from_selectors:
                try:
                    from_field = self.driver.find_element(By.XPATH, selector)
                    if from_field.is_displayed():
                        print(f"✓ Found From date field using: {selector}")
                        break
                except:
                    continue
            
            # Find To field  
            to_selectors = [
                "//span[text()='To:']/following::input[1]",
                "//label[contains(text(), 'To')]/following::input[1]",
                "//*[text()='To:']/following-sibling::*//input",
                "//input[contains(@id, 'to')]",
                "//input[contains(@name, 'to')]",
                "//td[text()='To:']/following-sibling::td//input"
            ]
            
            print("Searching for TO field...")
            for selector in to_selectors:
                try:
                    to_field = self.driver.find_element(By.XPATH, selector)
                    if to_field.is_displayed():
                        print(f"✓ Found To date field using: {selector}")
                        break
                except:
                    continue
            
            # FALLBACK: If TO field not found by label, try finding all date inputs
            if not to_field:
                print("\n⚠️ TO field not found by label, trying fallback method...")
                print("Looking for all date input fields...")
                
                try:
                    # Find all input fields that might be date fields
                    all_inputs = self.driver.find_elements(By.XPATH, "//input[@type='text' or not(@type)]")
                    date_inputs = []
                    
                    for inp in all_inputs:
                        if inp.is_displayed():
                            # Check if it looks like a date field
                            inp_value = inp.get_attribute('value') or ''
                            inp_id = inp.get_attribute('id') or ''
                            inp_name = inp.get_attribute('name') or ''
                            inp_class = inp.get_attribute('class') or ''
                            
                            # Look for date-related attributes or values
                            if (('/' in inp_value and len(inp_value) >= 8) or 
                                'date' in inp_id.lower() or 
                                'date' in inp_name.lower() or
                                'date' in inp_class.lower()):
                                date_inputs.append(inp)
                                print(f"  Found potential date field: value='{inp_value}', id='{inp_id}'")
                    
                    print(f"\n✓ Found {len(date_inputs)} potential date input fields")
                    
                    # If we found exactly 2 date inputs, assume first is FROM, second is TO
                    if len(date_inputs) >= 2:
                        from_field = date_inputs[0]
                        to_field = date_inputs[1]
                        print(f"✓ Using first field as FROM, second as TO")
                        print(f"  FROM field value: {from_field.get_attribute('value')}")
                        print(f"  TO field value: {to_field.get_attribute('value')}")
                    elif len(date_inputs) == 1:
                        print("⚠️ Only found 1 date field - this may not work")
                        # If we already have from_field, this must be TO
                        if from_field:
                            to_field = date_inputs[0]
                            print("  Assuming single field is TO field")
                    
                except Exception as e:
                    print(f"❌ Fallback method failed: {e}")
            
            # Another FALLBACK: Try finding TO field relative to FROM field
            if from_field and not to_field:
                print("\n⚠️ Trying to find TO field relative to FROM field...")
                try:
                    # Find the parent row/container of FROM field
                    parent = from_field.find_element(By.XPATH, "../..")
                    
                    # Find all inputs within this parent or next siblings
                    nearby_inputs = parent.find_elements(By.XPATH, ".//following::input[@type='text' or not(@type)]")
                    
                    # The first visible one after FROM should be TO
                    for inp in nearby_inputs[:3]:  # Check first 3 inputs
                        if inp.is_displayed() and inp != from_field:
                            to_field = inp
                            print(f"✓ Found TO field relative to FROM")
                            print(f"  TO field value: {to_field.get_attribute('value')}")
                            break
                            
                except Exception as e:
                    print(f"Relative search failed: {e}")
            
            # CRITICAL: Set the dates using JavaScript - EXACT ORIGINAL METHOD
            if from_field and to_field:
                print(f"\n✓ Both date fields found!")
                print(f"Setting FROM date to {start_str}...")
                
                # Use JavaScript to set FROM date
                try:
                    # Set value using JavaScript - ORIGINAL SCRIPT'S METHOD
                    self.driver.execute_script(f"""
                        arguments[0].value = '{start_str}';
                        arguments[0].setAttribute('value', '{start_str}');
                    """, from_field)
                    
                    # Trigger all possible change events - ORIGINAL SCRIPT'S METHOD
                    self.driver.execute_script("""
                        var element = arguments[0];
                        var event1 = new Event('change', { bubbles: true });
                        var event2 = new Event('input', { bubbles: true });
                        var event3 = new Event('blur', { bubbles: true });
                        element.dispatchEvent(event1);
                        element.dispatchEvent(event2);
                        element.dispatchEvent(event3);
                    """, from_field)
                    
                    print(f"✓ FROM date set to: {start_str}")
                except Exception as e:
                    print(f"❌ JavaScript method failed for FROM: {e}")
                    return False
                
                # Wait for FROM date to be processed
                print("Waiting for FROM date to process...")
                time.sleep(3)
                
                # Now set TO date - ORIGINAL SCRIPT'S METHOD
                print(f"\nSetting TO date to {end_str}...")
                try:
                    self.driver.execute_script(f"""
                        arguments[0].value = '{end_str}';
                        arguments[0].setAttribute('value', '{end_str}');
                    """, to_field)
                    
                    # Trigger change events
                    self.driver.execute_script("""
                        var element = arguments[0];
                        var event1 = new Event('change', { bubbles: true });
                        var event2 = new Event('input', { bubbles: true });
                        var event3 = new Event('blur', { bubbles: true });
                        element.dispatchEvent(event1);
                        element.dispatchEvent(event2);
                        element.dispatchEvent(event3);
                    """, to_field)
                    
                    print(f"✓ TO date set to: {end_str}")
                except Exception as e:
                    print(f"❌ Could not set TO date: {e}")
                    return False
                
                # Final verification - ORIGINAL SCRIPT'S METHOD
                time.sleep(2)
                try:
                    from_value = from_field.get_attribute('value')
                    to_value = to_field.get_attribute('value')
                    print(f"\n✓ Final verification - FROM: {from_value}, TO: {to_value}")
                except:
                    print("⚠️ Could not verify final date values")
                
                # Wait for dates to be applied and data to refresh - ORIGINAL TIMING
                print("\nWaiting for data to refresh with new date range...")
                time.sleep(5)
                
                return True
                
            else:
                print("\n❌ ERROR: Could not find BOTH date fields!")
                print(f"From field found: {from_field is not None}")
                print(f"To field found: {to_field is not None}")
                
                if from_field:
                    print(f"FROM field details:")
                    print(f"  - ID: {from_field.get_attribute('id')}")
                    print(f"  - Name: {from_field.get_attribute('name')}")
                    print(f"  - Value: {from_field.get_attribute('value')}")
                
                return False
                
        except Exception as e:
            print(f"❌ Failed to set date range: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def export_events(self, device_name, start_date, end_date):
        """Export events and convert to Excel - USING ORIGINAL WORKING LOGIC"""
        
        try:
            print("\n" + "="*50)
            print("Step 7: Cleaning up old files...")
            print("="*50)
            
            # Remove old files
            for ext in ['*.xlsx', '*.xls', '*.xml', '*.csv']:
                pattern = os.path.join(self.download_dir, f"*{device_name}*{ext}")
                old_files = glob.glob(pattern)
                for old_file in old_files:
                    try:
                        os.remove(old_file)
                        print(f"Removed old file: {os.path.basename(old_file)}")
                    except:
                        pass
            
            # Also remove any generic files that might conflict
            for ext in ['*.xml', '*.xls', '*.xlsx', '*.csv']:
                pattern = os.path.join(self.download_dir, ext)
                old_files = glob.glob(pattern)
                for old_file in old_files:
                    try:
                        os.remove(old_file)
                        print(f"Removed old file: {os.path.basename(old_file)}")
                    except:
                        pass
            
            print("\n" + "="*50)
            print("Step 8: Clicking Export Events button...")
            print("="*50)
            
            # The Export Events button - try ALL original selectors
            export_selectors = [
                "//span[text()='Export Events']/..",
                "//button[contains(., 'Export Events')]",
                "//*[contains(text(), 'Export Events')]",
                "//a[contains(text(), 'Export Events')]",
                "//div[@id='export' or contains(@class, 'export')]//span[contains(text(), 'Export')]/..",
                "//table[contains(@id, 'export')]//button"
            ]
            
            export_clicked = False
            print("Searching for Export Events button...")
            for selector in export_selectors:
                try:
                    export_elements = self.driver.find_elements(By.XPATH, selector)
                    print(f"Selector '{selector}' found {len(export_elements)} elements")
                    for export_btn in export_elements:
                        try:
                            if export_btn.is_displayed():
                                print(f"Found visible Export Events button")
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", export_btn)
                                time.sleep(0.5)
                                
                                try:
                                    export_btn.click()
                                    print("✓ Clicked with regular click")
                                except:
                                    self.driver.execute_script("arguments[0].click();", export_btn)
                                    print("✓ Clicked with JavaScript")
                                
                                export_clicked = True
                                print("✓ Clicked Export Events button!")
                                break
                        except:
                            continue
                    if export_clicked:
                        break
                except:
                    continue
            
            if not export_clicked:
                print("❌ ERROR: Could not click Export button")
                return None
            
            print("\nWaiting for download to complete...")
            time.sleep(10)
            
            # Check for downloaded file and convert to Excel
            print("\nChecking downloaded file and converting to Excel...")
            
            # Look for any new file in the download directory
            downloaded_files = []
            for ext in ['*.xml', '*.xls', '*.xlsx', '*.csv']:
                pattern = os.path.join(self.download_dir, ext)
                files = glob.glob(pattern)
                downloaded_files.extend(files)
            
            print(f"Found {len(downloaded_files)} potential download files")
            
            # Sort by modification time to get the newest file
            if downloaded_files:
                downloaded_files.sort(key=os.path.getmtime, reverse=True)
                newest_file = downloaded_files[0]
                print(f"Newest file: {newest_file}")
                print(f"Modified: {time.ctime(os.path.getmtime(newest_file))}")
                
                # Create base filename for converted files
                start_str = start_date.strftime("%Y-%m-%d")
                end_str = end_date.strftime("%Y-%m-%d")
                base_filename = f"{device_name}_Events_{start_str}_to_{end_str}"
                
                # Always try to convert to Excel, regardless of original extension
                excel_filename = os.path.join(self.download_dir, f"{base_filename}.xlsx")
                
                print(f"\nAttempting to convert to Excel format...")
                print(f"Target: {excel_filename}")
                conversion_success = self.convert_xml_to_excel(newest_file, excel_filename)
                
                if conversion_success:
                    print(f"✅ CONVERSION SUCCESSFUL!")
                    print(f"📊 Excel file created: {excel_filename}")
                    
                    # Keep original file as backup with descriptive name
                    file_ext = os.path.splitext(newest_file)[1].lower()
                    backup_filename = os.path.join(self.download_dir, f"{base_filename}_original{file_ext}")
                    
                    try:
                        if os.path.exists(backup_filename):
                            os.remove(backup_filename)
                        shutil.move(newest_file, backup_filename)
                        print(f"📁 Original file backed up as: {backup_filename}")
                    except:
                        print(f"⚠️ Could not backup original file, but Excel conversion succeeded")
                    
                    # File info
                    file_size = os.path.getsize(excel_filename)
                    print(f"📊 Excel file size: {file_size:,} bytes")
                    print(f"📅 Date range: {start_str} to {end_str}")
                    print(f"🎯 Ready for use!")
                    
                    return excel_filename
                else:
                    print(f"❌ CONVERSION FAILED - keeping original file")
                    # Just rename the original file with descriptive name
                    file_ext = os.path.splitext(newest_file)[1].lower()
                    final_filename = os.path.join(self.download_dir, f"{base_filename}{file_ext}")
                    
                    try:
                        if os.path.exists(final_filename):
                            os.remove(final_filename)
                        shutil.move(newest_file, final_filename)
                        print(f"📁 Original file saved as: {final_filename}")
                        print(f"⚠️ You may need to manually convert this file")
                        return final_filename
                    except Exception as e:
                        print(f"Could not rename file: {e}")
                        print(f"File available at: {newest_file}")
                        return newest_file
                        
            else:
                print("❌ WARNING: Could not find downloaded file")
                print("Contents of download directory:")
                try:
                    all_files = os.listdir(self.download_dir)
                    for f in all_files:
                        print(f"  - {f}")
                except:
                    pass
                return None
                
        except Exception as e:
            print(f"❌ Failed to export events: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def download_device_data(self, device_id, start_date, end_date, device_name=None):
        """
        Main method to download device data
        
        Args:
            device_id (str): Device ID to download data for
            start_date (datetime): Start date for data
            end_date (datetime): End date for data
            device_name (str): Device name (optional, will use ID if not provided)
        
        Returns:
            str: Path to downloaded Excel file, or None if failed
        """
        
        if device_name is None:
            device_name = f"Device-{device_id}"
        
        try:
            # Setup browser
            print("Setting up browser...")
            self.setup_driver()
            
            # Login
            if not self.login():
                return None
            
            # Navigate to device
            if not self.navigate_to_device(device_id, device_name):
                return None
            
            # Navigate to events
            if not self.navigate_to_events():
                return None
            
            # Set date range
            if not self.set_custom_date_range(start_date, end_date):
                return None
            
            # Export and convert
            excel_file = self.export_events(device_name, start_date, end_date)
            
            if excel_file:
                print(f"\n✅ SUCCESS: Downloaded and converted to {excel_file}")
            
            return excel_file
            
        except Exception as e:
            print(f"✗ Error during download process: {e}")
            import traceback
            traceback.print_exc()
            return None
            
        finally:
            # Close browser
            if self.driver:
                time.sleep(2)
                self.driver.quit()
                print("Browser closed")


# For testing independently
if __name__ == "__main__":
    downloader = UplinkDownloader(headless=False)
    
    test_device_id = "359205108536865"
    test_device_name = "Scout-12197"
    test_start = datetime(2025, 6, 10)  # Commission date
    test_end = datetime.now()
    
    excel_file = downloader.download_device_data(
        device_id=test_device_id,
        start_date=test_start,
        end_date=test_end,
        device_name=test_device_name
    )
    
    if excel_file:
        print(f"\n✓ SUCCESS: Downloaded file to {excel_file}")
    else:
        print("\n✗ FAILED: Could not download file")