#!/usr/bin/env python
# coding: utf-8

# In[17]:


import os
import openpyxl
from selenium import webdriver
from bs4 import BeautifulSoup
from tqdm import tqdm

# Function to scrape agent details
def scrape_agent_details(agent_url):
    driver = webdriver.Chrome()
    driver.get(agent_url)
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, "html.parser")
    
    name_element = soup.find("p", class_="profile-Tiltle-main")
    name = name_element.get_text(strip=True) if name_element else "N/A"
    
    office_number_span = soup.find("span", class_="mobile-number", string=" Office")
    office_number = office_number_span.find_previous("span").get_text(strip=True) if office_number_span else "N/A"
    
    mobile_number_span = soup.find("span", class_="mobile-number", string=" Mobile")
    mobile_number = mobile_number_span.find_previous("span").get_text(strip=True) if mobile_number_span else "N/A"
    
    company_element = soup.find("p", class_="jsx-3831114755 addressspace")
    company_name = company_element.get_text(strip=True) if company_element else "N/A"
    
    address_elements = soup.select("div.jsx-3831114755.better-homes-and-gar-icon-right p.agent_address span")
    address_parts = [element.get_text(strip=True) for element in address_elements]
    full_address = ", ".join(address_parts) if address_parts else "N/A"
    
    driver.quit()
    
    return name, office_number, mobile_number, company_name, full_address

# Main function
def main():
    # URL of the initial page
    url = "https://www.realtor.com/realestateagents/atlanta_ga/pg-11"
    
    # Initialize the WebDriver
    driver = webdriver.Chrome()
    driver.get(url)
    
    # Get the page source after JavaScript execution
    page_source = driver.page_source
    
    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(page_source, "html.parser")
    
    # Find all agent card divs
    agent_cards = soup.find_all("div", class_="jsx-2987058905 agent-list-card clearfix")
    
    # Extract and store agent URLs
    base_url = "https://www.realtor.com"
    agent_urls = [base_url + card.find("a", class_="jsx-2987058905")["href"] for card in agent_cards]
    
    # Close the WebDriver
    driver.quit()
    
    # Output file path
    output_file_path = "D:\\Office\\Scripts\\Agent\\agent_info_11.xlsx"
    
    # Check if the output file already exists
    if os.path.exists(output_file_path):
        print("Output file already exists. Please remove or rename the existing file.")
        return
    
    # Create a new Excel workbook and add a worksheet
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    
    # Write header row to the worksheet
    header = ["Name", "Office", "Mobile", "Company", "Address"]
    worksheet.append(header)
    
    # Total number of agents
    total_agents = len(agent_urls)

    # Scrape agent details and save to Excel
    for idx, agent_url in enumerate(agent_urls, start=1):
        agent_data = scrape_agent_details(agent_url)

        print(f"\nAgent Information {idx} out of {total_agents}")
        print("Agent Url:", agent_url)
        print("Name:", agent_data[0])
        print("Office:", agent_data[1])
        print("Mobile:", agent_data[2])
        print("Company:", agent_data[3])
        print("Address:", agent_data[4])
        print("-" * 30)

        worksheet.append(agent_data)

    # Save the Excel workbook
    workbook.save(output_file_path)
    print("Agent information saved to:", output_file_path)

if __name__ == "__main__":
    main()


# In[ ]:




