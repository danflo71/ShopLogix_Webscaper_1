from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import shutil
import xlwings as xw
import logging
import os
import time
import schedule
import pandas as pd

# Setting up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def take_screenshot(driver, name):
    screenshot_path = f'screenshots/{name}.png'
    driver.save_screenshot(screenshot_path)
    logger.info(f'Screenshot saved to {screenshot_path}')

def wait_for_file(file_path, timeout=30):
    start_time = time.time()
    while True:
        if os.path.exists(file_path):
            logger.info(f'File found: {file_path}')
            return True
        elif time.time() - start_time > timeout:
            logger.error(f'Timeout reached: {file_path} not found')
            return False
        time.sleep(1)

# PM Criticality Script
def update_Criticality():
    driver = None
    try:
        logger.info('Starting update process')

        # Define Edge options to use existing user profile
        edge_options = Options()
        edge_options.add_argument("--no-sandbox")
        edge_options.add_argument("--disable-dev-shm-usage")
        edge_options.add_argument("--headless=new")  # Enable headless mode
        edge_options.add_argument("--disable-gpu")  # Disable GPU acceleration
       
        # Path to your Edge user data directory and profile
        user_data_dir = "C:/Users/COMPUTERUSERNAME/AppData/Local/Microsoft/Edge/User Data"   # (This will change depending on you computer)
        profile_dir = "Default"  # Replace 'Default' with your profile name if different
        edge_options.add_argument(f"user-data-dir={user_data_dir}")
        edge_options.add_argument(f"profile-directory={profile_dir}")

        # Clear the file that we are looking for inside of Downloads folder
        original_data_file_path = "C:/Users/COMPUTERUSERNAME/Downloads/BDGEHOURS.csv"   #(This will change depending on wher you are downloading the file depending on your computer)
        if os.path.exists(original_data_file_path):
            os.remove(original_data_file_path)
            logger.info('Original file deleted')
         
        # Locates the Excel sheet from the teams folder
        teams_excel_file_path = "C:/Users/COMPUTERUSERNAME/OneDrive - LOCATION/FileLocation/FileName.xlsx" #(The front part of the path will change but not the seoncd part ex C:/Users/COMPUTERUSERNAME/OneDrive - LOCATION/ )
       
           
        # Initialize Edge WebDriver
        service = EdgeService(executable_path='C:/Users/COMPUTERUSERNAME/OneDrive - LOCATION/VS_Code_Project_FIle_name/drivers/msedgedriver.exe') # ( This driver will have to be changed to your driectory for the driver)
        driver = webdriver.Edge(service=service, options=edge_options)
       
        logger.info('WebDriver initialized')

        # Create screenshots directory if it doesn't exist
        if not os.path.exists('screenshots'):
            os.makedirs('screenshots')

        # Navigate to the login page
        login_url = 'https://portal.shoplogix.com/Privatelink'
        driver.get(login_url)
        logger.info('Navigated to login page')
        take_screenshot(driver, 'login_page')

        # Increase the timeout time
        timeout = 100
   
        logger.info('Navigated to target URL')
        take_screenshot(driver, 'target_page')

        logger.info('Data Sheet located')
        time.sleep(4)
        driver.implicitly_wait(60)
       
        # Hover over the section where the element becomes visible
        hover_element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#prism-mainview > div.prism-persistent-mainview-holder.slf-back-secondary > div > div.collapsible-pane.errvil-warpper > div > dashboard > div > div.content > div > div:nth-child(1) > div > div > div:nth-child(2) > div:nth-child(1) > div > widget > widget-header > div.widget-title__holder.widget-title__holder--view.uneditable-title > widget-title'))
        )
       
        # the section above inside of the ' ' changes based on the item you are trying to click on
       
        take_screenshot(driver, 'hover_page')
        time.sleep(4)
        ActionChains(driver).move_to_element(hover_element).perform()
        logger.info('Hovered over the target section')

        threeDots = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="prism-mainview"]/div[2]/div/div[3]/div/dashboard/div/div[1]/div/div[1]/div/div/div[2]/div[1]/div/widget/widget-header/widget-toolbar/div[2]/button[7]/span'))
        )
        # the section above inside of the ' ' changes based on the item you are trying to click on
        threeDots.click()
        take_screenshot(driver, 'click_3dots')
        logger.info('Clicked 3 Dots Button')
       
        # Locate the download tab
        download_tab = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > data-menu > div > div > div > div > div:nth-child(1) > div > div.menu-item-arrow-holder > span'))
        )
        # the section above inside of the ' ' changes based on the item you are trying to click on
        take_screenshot(driver, 'download_click')
        download_tab.click()
        logger.info('Clicked the Download tab')
       
        # Locate the download CSV option
        download_csv = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > data-menu > div > div:nth-child(2) > div > div > div:nth-child(3) > div > div.mi-caption'))
        )
        # the section above inside of the ' ' changes based on the item you are trying to click on
        take_screenshot(driver, 'csv_click')
        download_csv.click()
       
        time.sleep(10)
        logger.info('Clicked download_csv')
        take_screenshot(driver, 'after_clicking_download_csv')

        new_file_path = "C:/Users/COMPUTERUSERNAME/Downloads/FILEName2.csv" # This file path will change  ( "C:/Users/COMPUTERUSERNAME/Downloads/ )
       
        if not wait_for_file(new_file_path):
            raise FileNotFoundError(f"Downloaded file not found: {new_file_path}")
       
        logger.info(f'Downloaded file found: {new_file_path}')
       
        # Load the new data from the CSV file
       
        df = pd.read_csv(new_file_path)  # Adjust the date column name as needed

        # Open the Excel file and update it with the new data using xlwings
        wb = xw.Book(teams_excel_file_path)
        ws = wb.sheets['BDGEHOURS']  # Ensure this matches the sheet name in your Excel file
       
        # Clear existing data in columns A to I
        ws.range('A1:I' + str(ws.cells.last_cell.row)).clear_contents()
       
        # Write the new data to the sheet
        ws.range('A1').options(index=False, header=True).value = df
       
        # Save the updated Excel file
        wb.save(teams_excel_file_path)
        wb.close()
       
        logger.info('Excel file updated and data is now ready')

    except Exception as e:
        logger.error(f'Error in update_Criticality {e}')
        if driver:
            take_screenshot(driver, 'error')
    finally:
        if driver:
            driver.quit()
            logger.info('Driver closed')    
       
       
       
        logger.info('Finished PM Criticality Update')
       

def pm_criticality_job():
    update_Criticality()

# Update Excel Script
def update_excel():
    driver = None
    try:
        logger.info('Starting update process')

        # Define Edge options to use existing user profile
        edge_options = Options()
        edge_options.add_argument("--no-sandbox")
        edge_options.add_argument("--disable-dev-shm-usage")
        edge_options.add_argument("--headless=new")  # Enable headless mode
        edge_options.add_argument("--disable-gpu")  # Disable GPU acceleration
       
        # Path to your Edge user data directory and profile
        user_data_dir = "C:/Users/COMPUTERUSERNAME/AppData/Local/Microsoft/Edge/User Data"   # this changes based on where your data is stored it should follow the same path directory
        profile_dir = "Default"  # Replace 'Default' with your profile name if different
        edge_options.add_argument(f"user-data-dir={user_data_dir}")
        edge_options.add_argument(f"profile-directory={profile_dir}")

        # Set download preferences
       
        # Clear the file that we are looking for inside of Downloads folder
        original_file_path= "C:/Users/COMPUTERUSERNAME/Downloads/FileDownload2.csv"  # The only part that changes it the "C:/Users/COMPUTERUSERNAME/Downloads/ this will be set to your download folder in your user but what you  DONT change is the name of the file
        if os.path.exists(original_file_path):
            os.remove(original_file_path)
            logger.info('Original file deleted')
         
        # Removes the file sheet from the teams folder    
        original_teams_file_path= "C:/Users/COMPUTERUSERNAME/OneDrive - Location/NameofFile/FileDownload2.csv" # The only part that changes it the "C:/Users/COMPUTERUSERNAME/OneDrive - MDLZ/FileName/ this will be set to your download folder in your user but what you  DONT change is the name of the file
        if os.path.exists(original_teams_file_path):
            os.remove(original_teams_file_path)
            logger.info('Original file deleted')
       
   

        # Initialize Edge WebDriver
        service = EdgeService(executable_path='C:\\Users\\COMPUTERUSERNAME\\OneDrive - Location\\Documents\\VS_Code_Project_File_name\\drivers\\msedgedriver.exe') # This driver location will change depending on you locaiton that you downloaded it in make sure to use doulbe \\ do adress location
        driver = webdriver.Edge(service=service, options=edge_options)
       
       
        logger.info('WebDriver initialized')

        # Create screenshots directory if it doesn't exist
        if not os.path.exists('screenshots'):
            os.makedirs('screenshots')

        # Navigate to the login page
        login_url = 'https://portal.shoplogix.com/PrivateLink' #link DOES NOT CHANGE unless the link of the page changes which hassnt happen
        driver.get(login_url)
        logger.info('Navigated to login page')
        take_screenshot(driver, 'login_page')

        # Increase the timeout time
        timeout = 100
   
        logger.info('Navigated to target URL')
        take_screenshot(driver, 'target_page')

        logger.info('Data Sheet located')
        time.sleep(4)
        driver.implicitly_wait(60)
       
        # Hover over the section where the element becomes visible
        hover_element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '#prism-mainview > div.prism-persistent-mainview-holder.slf-back-secondary > div > div.collapsible-pane.errvil-warpper > div > dashboard > div > div.content > div > div:nth-child(1) > div > div > div > div.dashboard-layout-subcell-host > div > widget > widget-header > div.widget-title__holder.widget-title__holder--view.uneditable-title > widget-title'))
        )
        # the section above inside of the ' ' changes based on the item you are trying to click on
        take_screenshot(driver, 'hover_page')
        time.sleep(4)
        ActionChains(driver).move_to_element(hover_element).perform()
        logger.info('Hovered over the target section')

        threeDots = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="prism-mainview"]/div[2]/div/div[3]/div/dashboard/div/div[1]/div/div[1]/div/div/div/div[1]/div/widget/widget-header/widget-toolbar/div[2]/button[7]'))
        )
        # the section above inside of the ' ' changes based on the item you are trying to click on
        threeDots.click()
        take_screenshot(driver, 'click_3dots')
        logger.info('Clicked 3 Dots Button')
       
        # Locate the download tab
        download_tab = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > data-menu > div > div:nth-child(1) > div > div > div:nth-child(1) > div > div.mi-caption'))
        )
        # the section above inside of the ' ' changes based on the item you are trying to click on
        take_screenshot(driver, 'download_click')
        download_tab.click()
        logger.info('Clicked the Download tab')
       
        # Locate the download CSV option
        download_csv = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'body > data-menu > div > div:nth-child(2) > div > div > div:nth-child(3) > div > div.mi-caption'))
        )
        # the section above inside of the ' ' changes based on the item you are trying to click on
        take_screenshot(driver, 'csv_click')
        download_csv.click()
       
        time.sleep(2)
        logger.info('Clicked download_csv')
        take_screenshot(driver, 'after_clicking_download_csv')

        new_file_path = "C:/Users/COMPUTERUSERNAME/Downloads/FileDownload2.csv" # This section of the code ischanged the only section that should be changed is "C:/Users/COMPUTERUSERNAME/Downloads/ It should point to you file where you have it downloaded
       
        if not wait_for_file(new_file_path):
            raise FileNotFoundError(f"Downloaded file not found: {new_file_path}")
       
        logger.info(f'Downloaded file found: {new_file_path}')
       
        new_destination= "C:/Users/COMPUTERUSERNAME/OneDrive - MDLZ/FileName/FileDownload2.csv" # This section whilll change but only the pointer to the location of the file "C:/Users/COMPUTERUSERNAME/OneDrive - MDLZ/
        shutil.move(new_file_path, new_destination)
       
        # Process the downloaded file here
        driver.implicitly_wait(10)
        # Goes to file location and changes file name

        logger.info('File downloaded and successfully overwritten old file')

    except Exception as e:
        logger.error(f'Error in update_excel: {e}')
        if driver:
            take_screenshot(driver, 'error')
    finally:
        if driver:
            driver.quit()
            logger.info('Driver closed')
        logger.info('Finished Update Excel')

def update_excel_job():
    update_excel()

# Schedule the jobs
schedule.every().hour.at(":58").do(pm_criticality_job)
schedule.every().hour.at(":00").do(update_excel_job)

# Infinite loop to run the scheduled tasks
while True:
    schedule.run_pending()
    time.sleep(1)
