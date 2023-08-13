#!/usr/bin/env python
# coding: utf-8

# In[4]:


import os
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm  # for progress bar
import pandas as pd

base_url = "https://www.century21.com"
start_page = 21
end_page = 40

agent_info_list = []

for page in tqdm(range(start_page, end_page + 1), desc="Scraping progress"):
    page_url = base_url + f"/real-estate-agents/south-carolina/LSSC/?s={(page - 1) * 12}"
    
    response = requests.get(page_url)
    soup = BeautifulSoup(response.content, 'html.parser')

    for link in soup.find_all('a', class_='stretched-link'):
        href = link.get('href')
        if "/real-estate-agent/profile/" in href:
            full_agent_link = base_url + href

            agent_response = requests.get(full_agent_link)
            agent_soup = BeautifulSoup(agent_response.content, 'html.parser')

            # Extract agent information
            name = agent_soup.find('h1', class_='h2').strong.get_text()

            company_element = agent_soup.find('span', class_='lead d-block')
            company = company_element.get_text(strip=True).replace('of ', '') if company_element else "N/A"

            try:
                address_element = agent_soup.find('address')
                address = address_element.get_text(strip=True).replace('\n', ' ') if address_element else "N/A"
            except AttributeError:
                address = "N/A"

            try:
                mobile_div = agent_soup.find('p', class_='lh-lg')
                mobile_text = mobile_div.get_text(strip=True) if mobile_div else "N/A"

            except AttributeError:
                mobile_text = "N/A"

            # Add agent information to the list
            agent_info = {
                'Name': name,
                'Company': company,
                'Address': address,
                'Mobile': mobile_text,
                'Office': "N/A"  # No need for the office cell
            }
            agent_info_list.append(agent_info)

# Create a DataFrame from the collected agent information
df = pd.DataFrame(agent_info_list)

# Declare the output file path
output_file_path = r'D:\\Office\\Scripts\\south_carolina_2.xlsx'

# Check if the output file already exists
if os.path.exists(output_file_path):
    print("Output file already exists. Please remove or rename the existing file.")
else:
    # Save the DataFrame to an Excel file
    df.to_excel(output_file_path, index=False)
    print("Data saved to", output_file_path)
    print("Date Scraped", len(df))


# In[ ]:




