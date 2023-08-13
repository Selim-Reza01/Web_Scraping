#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Set up the Selenium WebDriver
driver = webdriver.Chrome()  # Make sure you have the Chrome driver in your PATH

# Define the file path to read URLs from
input_file_path = "D:\\Office\\Scripts\\input.xlsx"

# Read the Excel file to get the list of URLs
url_df = pd.read_excel(input_file_path)
urls = url_df['url'].tolist()  # Change 'Page URL' to the actual column name

# Loop through each URL
for idx, url in enumerate(urls, start=1):
    print("-"*30)
    print(f"Extracting {idx} out of {len(urls)}")
    
    def has_captcha_page(url):
        driver.get(url)
        captcha_meta_tag = driver.find_elements(By.XPATH, "//head/meta[@name='ROBOTS'][@content='NOINDEX, NOFOLLOW']")
        return captcha_meta_tag != []

    def solve_captcha_manually():
        print("Additional security check is required. Please manually solve the CAPTCHA.")
        time.sleep(15)

    # Load the URL and check for CAPTCHA
    driver.get(url)
    if has_captcha_page(url):
        solve_captcha_manually()

    # Reload the page if CAPTCHA is still present
    reload_attempts = 1
    for _ in range(reload_attempts):
        driver.refresh()
        if has_captcha_page(url):
            solve_captcha_manually()
        else:
            break

    try:
        # Continue with data collection
        wait = WebDriverWait(driver, 10)

        price_element = wait.until(EC.presence_of_element_located((By.ID, "listingPriceValue")))
        price = price_element.text.strip() if price_element else "N/A"

        address_element = wait.until(EC.presence_of_element_located((By.ID, "listingAddress")))
        address = address_element.text.strip().replace('<br>', ', ') if address_element else "N/A"
        address = address.replace('\n', ', ')

        agent_name_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "realtorCardName")))
        agent_name = agent_name_element.text.strip() if agent_name_element else "N/A"

        brokerage_name_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "officeCardName")))
        brokerage_name = brokerage_name_element.text.strip() if brokerage_name_element else "N/A"

        iframe_element = wait.until(EC.presence_of_element_located((By.ID, "NeighbourhoodiFrame")))
        iframe_data_src = iframe_element.get_attribute("data-src")

        latitude = None
        longitude = None

        if iframe_data_src:
            data_src_parts = iframe_data_src.split('?')[1].split('&')
            for part in data_src_parts:
                if part.startswith('Latitude='):
                    latitude = part.replace('Latitude=', '')
                elif part.startswith('Longitude='):
                    longitude = part.replace('Longitude=', '')

        # Create a dictionary with the collected data
        data = {
            "Link": url,
            "Price": price,
            "Address": address,
            "Agent Name": agent_name,
            "Brokerage Name": brokerage_name,
            "Latitude": latitude if latitude else "N/A",
            "Longitude": longitude if longitude else "N/A"
        }

        # Convert the dictionary to a DataFrame
        df = pd.DataFrame([data])

        # Define the output file path to save the Excel file
        output_file_path = "D:\\Office\\Scripts\\data_240.xlsx"

        # Check if the Excel file exists
        if os.path.exists(output_file_path):
            # Append to the existing Excel file
            existing_df = pd.read_excel(output_file_path)
            df = pd.concat([existing_df, df], ignore_index=True)

        # Save the DataFrame to an Excel file
        df.to_excel(output_file_path, index=False)
        
        # Print success message
        print("Data scraped.")
        
    except Exception as e:
        print("Find Error. Skipping this!")
        
print("\nCongratulations!! Scraping Complete. Data Saved to ",output_file_path)

# Close the browser after processing all URLs
driver.quit()


# In[ ]:




