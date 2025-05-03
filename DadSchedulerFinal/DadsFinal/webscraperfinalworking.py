from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time
import os
from datetime import datetime
import re

def setup_driver():
    """Set up the Chrome WebDriver with appropriate options"""
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")  # Start with maximized browser
    
    # Optional: Uncomment to run in headless mode (no browser UI)
    # chrome_options.add_argument("--headless")
    
    # Set up driver - update the path to your chromedriver if needed
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def login_to_website(driver, url, username, password):
    """Login to the website with provided credentials"""
    print(f"Opening website: {url}")
    driver.get(url)
    
    # Wait for the page to load
    time.sleep(2)
    
    try:
        # Try different possible selectors for username field
        username_selectors = [
            #(By.ID, "User ID"),
            #(By.NAME, "UserName"),
            #(By.NAME, "userid"),
            (By.XPATH, "//input[@type='text' and contains(@name, 'UserName')]"),
            (By.XPATH, "//input[@name='UserName' and contains(@id, 'UserName')]")
        ]
        
        username_field = None
        for selector in username_selectors:
            try:
                username_field = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(selector)
                )
                if username_field:
                    break
            except:
                continue
                
        if not username_field:
            raise NoSuchElementException("Could not find username field")
            
        username_field.clear()
        username_field.send_keys(username)
        print("Username entered")
        
        # Try different possible selectors for password field
        password_selectors = [
            #(By.ID, "Password"),
            (By.NAME, "Password"),
            (By.XPATH, "//input[@type='password']")
        ]
        
        password_field = None
        for selector in password_selectors:
            try:
                password_field = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(selector)
                )
                if password_field:
                    break
            except:
                continue
                
        if not password_field:
            raise NoSuchElementException("Could not find password field")
            
        password_field.clear()
        password_field.send_keys(password)
        print("Password entered")
        
        # Try different possible selectors for login button
        login_button_selectors = [
            (By.XPATH, "//input[@type='submit']"),
            #(By.XPATH, "//button[contains(text(), 'Login')]"),
            (By.XPATH, "//input[@name='Submit']"),
            #(By.XPATH, "//button[@type='submit']")
        ]
        
        login_button = None
        for selector in login_button_selectors:
            try:
                login_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(selector)
                )
                if login_button:
                    break
            except:
                continue
                
        if not login_button:
            raise NoSuchElementException("Could not find login button")
            
        login_button.click()
        print("Login button clicked")
        
        # Wait for successful login - try different possible success indicators
        success_indicators = [
            (By.XPATH, "//td[contains(text(), 'ITMS Vendor Portal')]"),
            (By.XPATH, "//span[contains(text(), 'Welcome REMT !')]"),
            (By.XPATH, "//span[@id='welcomeUser']"),
            #(By.XPATH, "//div[contains(text(), 'Logout')]")
        ]
        
        success = False
        for indicator in success_indicators:
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(indicator)
                )
                success = True
                break
            except:
                continue
                
        if not success:
            raise TimeoutException("Could not verify successful login")
            
        print("Successfully logged in")
        return True
        
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error during login: {e}")
        # Take a screenshot for debugging
        driver.save_screenshot("login_error.png")
        print("Screenshot saved as login_error.png")
        return False

def click_red_car_icon(driver):
    """Click on the red car picture to download data - visible in Image 2"""
    try:
        # From Image 2, there's a red car icon with text "Download today's Trips/Set Rates/Generate Invoices"
        red_car = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//img[contains(@src, 'images/DownloadTrips.gif')]"))
        )
        # Alternative XPATH based on the surrounding text
        if not red_car:
            red_car = driver.find_element(By.XPATH, "//*[contains(text(), 'Download today's Trips/Set Rates/Generate Invoices')]")
        
        red_car.click()
        print("Clicked on red car icon")
        
        # Wait for data page to load - based on Image 3, we're looking for the "MH Trips" tab
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//img[contains(@src, '/ITMSVP/images/tabTripsOnMART.gif')]"))
        )
        
        # Make sure we're on the Trips tab (Image 3 shows we need to be on MH Trips)
        trips_tab = driver.find_element(By.XPATH, "//img[contains(@src, '/ITMSVP/images/tabTripsOnMART.gif')]")
        if '/ITMSVP/images/tabTripsOnMART.gif' in trips_tab.text:
            print("Already on MH Trips tab")
        else:
            trips_tab.click()
            print("Clicked on MH Trips tab")
            time.sleep(2)
        
        return True
        
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error clicking red car icon: {e}")
        return False

def scrape_table_data(driver):
    """Scrape data from the current page by selecting all rows"""
    all_data = []
    
    try:
        # Wait for the table to be visible
        print("Waiting for table to load...")
        table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "grid_MainDataGrid"))
        )
        print("Table found")
        
        # Get the table headers
        print("Getting table headers...")
        headers = []
        # Get all header cells using the provided XPath
        header_cells = driver.find_elements(By.XPATH, "//*[@id='grid_MainDataGrid']/tbody/tr[1]/td")
        for cell in header_cells:
            headers.append(cell.text.strip())
        
        print(f"Found {len(headers)} headers: {headers}")
        
        # Add headers as first row if they exist
        if headers:
            all_data.append(headers)
        
        # Get all data rows from the table
        print("Getting table rows...")
        # Get all rows except the header row
        rows = driver.find_elements(By.XPATH, "//*[@id='grid_MainDataGrid']/tbody/tr[position()>1]")
        print(f"Found {len(rows)} rows")
        
        for row in rows:
            # Get all cells in the row, including empty ones
            cells = row.find_elements(By.XPATH, ".//td")
            row_data = []
            for cell in cells:
                # Get the cell's class to identify its type
                cell_class = cell.get_attribute("class")
                # Check if cell is part of the data structure we want
                if any(cls in cell_class for cls in ["DataGrid-ItemStyle-ControlColumn", 
                                                   "aspNetDisabled DataGrid-ItemStyle-ControlColumn", 
                                                   "DataGrid-ItemStyle"]):
                    row_data.append(cell.text.strip())
            
            # Add the row data, even if empty, to maintain table structure
            all_data.append(row_data)
        
        print(f"Scraped {len(all_data) - 1} rows of data")  # Subtract 1 for header row
        
        # Take a screenshot of the table for debugging
        driver.save_screenshot("table_scrape.png")
        print("Screenshot saved as table_scrape.png")
        
        return all_data
        
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error scraping data: {e}")
        # Take a screenshot when error occurs
        driver.save_screenshot("table_error.png")
        print("Error screenshot saved as table_error.png")
        return []

def go_to_next_page(driver):
    """Click on the next arrow to go to the next page"""
    try:
        # Wait for the page info to be visible
        page_info = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Page')]"))
        ).text
        
        current_page, total_pages = map(int, re.findall(r'\d+', page_info))
        
        if current_page >= total_pages:
            print("No more pages available")
            return False
            
        # Look for the next page button
        next_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//img[contains(@src, 'arwSmallDownOn.gif')]"))
        )
        
        next_button.click()
        print(f"Navigated to page {current_page + 1} of {total_pages}")
        time.sleep(2)  # Wait for page to load
        return True
        
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error navigating to next page: {e}")
        return False

def save_data_to_excel(all_data, filepath):
    """Save the collected data to an Excel file"""
    try:
        # Convert data to pandas DataFrame
        if all_data and len(all_data) > 1:  # Check if data exists and has at least one row plus headers
            headers = all_data[0]
            data = all_data[1:]
            
            # Remove any empty columns
            headers = [h for h in headers if h]
            data = [[cell for cell in row if cell] for row in data]
            
            # Create DataFrame
            df = pd.DataFrame(data)
            
            # If the DataFrame has the right number of columns, assign headers
            if len(headers) == len(df.columns):
                df.columns = headers
            
            # Save to Excel
            df.to_excel(filepath, index=False)
            print(f"Data saved to {filepath}")
            return True
        else:
            print("No data to save")
            return False
            
    except Exception as e:
        print(f"Error saving data to Excel: {e}")
        return False

def main():
    # Website credentials - from Image 1
    url = "enter url here"  # Update with your actual URL
    username = "enter username here"
    password = "enter password here"
    
    # Setup Chrome driver
    driver = setup_driver()
    
    try:
        # Login to website
        if not login_to_website(driver, url, username, password):
            print("Login failed. Exiting.")
            driver.quit()
            return
        
        # Click on red car icon from the main page (Image 2)
        if not click_red_car_icon(driver):
            print("Failed to access data through red car icon. Exiting.")
            driver.quit()
            return
        
        # Initialize all_data list to store data from all pages
        all_data = []
        
        # Handle pagination and scrape data
        has_more_pages = True
        page_count = 1
        
        while has_more_pages:
            print(f"Processing page {page_count}")
            
            # Scrape current page
            page_data = scrape_table_data(driver)
            
            # Add headers only from the first page
            if page_count == 1 and page_data:
                all_data.extend(page_data)
            elif page_data and len(page_data) > 1:
                # Skip headers for subsequent pages
                all_data.extend(page_data[1:])
            
            # Try to go to next page
            has_more_pages = go_to_next_page(driver)
            page_count += 1
        
        # Create directory if it doesn't exist
        output_dir = os.path.join(os.path.expanduser("~"), "Downloads", "WebScrapedData")
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate filename with current date
        current_date = datetime.now().strftime("%Y%m%d")
        excel_path = os.path.join(output_dir, f"MART_Trips_{current_date}.xlsx")
        
        # Save data to Excel
        save_data_to_excel(all_data, excel_path)
        
    except Exception as e:
        print(f"An error occurred: {e}")
    
    finally:
        # Close the browser
        print("Closing browser")
        driver.quit()

if __name__ == "__main__":
    main()