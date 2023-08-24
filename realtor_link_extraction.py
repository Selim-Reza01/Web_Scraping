#!/usr/bin/env python
# coding: utf-8

# In[14]:


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# Base URL
base_url = "https://www.realtor.ca/map#view=list&CurrentPage={page_number}&Sort=6-D&GeoIds=g30_f241etq5&GeoName=Ottawa%2C%20ON&PropertyTypeGroupID=1&PropertySearchTypeId=1&PriceMin=1500000&Currency=CAD"

# Set the start and end page numbers
start_page = 1
end_page = 24

# Initialize the WebDriver
driver = webdriver.Chrome()  # Make sure chromedriver is in your PATH

def has_captcha_page(url):
    driver.get(url)
    captcha_meta_tag = driver.find_elements(By.XPATH, "//head/meta[@name='ROBOTS'][@content='NOINDEX, NOFOLLOW']")
    return captcha_meta_tag != []

def solve_captcha_manually():
    print("Additional security check is required. Please manually solve the CAPTCHA.")
    time.sleep(15)

try:
    extracted_links = []  # To store extracted links
    
    for page_number in range(start_page, end_page + 1):
        # Construct the URL for the current page
        url = base_url.replace("{page_number}", str(page_number))
        
        # Load the URL and check for CAPTCHA
        print("-"*30)
        print("Loading page : ", page_number, "out of ",end_page )
    
        driver.get(url)
        if has_captcha_page(url):
            solve_captcha_manually()
            
            # Reload the page if CAPTCHA is still present
            driver.refresh()
            if has_captcha_page(url):
                solve_captcha_manually()
                continue  # Skip to the next iteration if CAPTCHA is detected
        
        # Explicitly wait for the listingDetailsLink elements to be present
        wait = WebDriverWait(driver, 10)
        link_elements = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "listingDetailsLink")))
        
        links_on_page = [link.get_attribute("href") for link in link_elements]
        extracted_links.extend(links_on_page)
    
    # Save all the extracted links to an Excel file
    excel_file_path = 'D:\\Office\\Scripts\\ottawa_data_set_5.xlsx'  # Set the desired file path
    data = {'Links': extracted_links}
    df = pd.DataFrame(data)
    df.to_excel(excel_file_path, index=False)
    
    # Print the path to the saved Excel file
    print("Congratulations !! Scraping Complete.Total data found",len(extracted_links),". File saved at:", excel_file_path)
    
finally:
    # Close the WebDriver at the end
    driver.quit()


# In[ ]:




