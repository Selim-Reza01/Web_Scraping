#!/usr/bin/env python
# coding: utf-8

# In[2]:


import requests
from bs4 import BeautifulSoup
import openpyxl
import os
from tqdm import tqdm
from datetime import datetime

# Function to scrape agent details from a single page
def scrape_page(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')

    agent_data = {}

    name_element = soup.find('h1', class_='agent-content-name')
    if name_element:
        a_element = name_element.find('a')
        agent_data['Name'] = a_element.text if a_element else ''
    else:
        name_span = soup.find('h1', id='main-content').find('span', class_='notranslate')
        agent_data['Name'] = name_span.text if name_span else ''

    role_element = soup.find('h2', itemprop='jobTitle')
    agent_data['Role'] = role_element.text if role_element else ''

    email_element = soup.find('a', class_='email-link')
    agent_data['Email'] = email_element.text if email_element else ''

    mobile_element = soup.find('a', {'class': 'phone-link', 'data-phone-type': 'mobile'})
    agent_data['Mobile'] = mobile_element.text if mobile_element else ''

    office_element = soup.find('a', {'class': 'phone-link', 'data-phone-type': 'office'})
    agent_data['Office'] = office_element.text if office_element else ''

    direct_element = soup.find('a', {'class': 'phone-link', 'data-phone-type': 'direct'})
    agent_data['Direct'] = direct_element.text if direct_element else ''

    market_element = soup.find('a', class_='line market-link')
    agent_data['Market'] = market_element.text if market_element else ''

    company_div = soup.find('div', class_='body notranslate')
    company_element = company_div.find('a') if company_div else None
    agent_data['Company'] = company_element.text if company_element else ''

    address_element = soup.find('span', class_='office-span')
    agent_data['Address'] = clean_address(address_element.text) if address_element else ''

    return agent_data

# Function to clean up the address text
def clean_address(address):
    lines = address.strip().split('\n')
    cleaned_address = '\n'.join(line.strip() for line in lines if line.strip())
    return cleaned_address


# Function to scrape data from a specific page
def scrape_specific_pages(base_url, start_page, end_page):
    all_agent_data = []

    for page_num in tqdm(range(start_page, end_page + 1), desc="Scraping progress", unit='page'):
        if page_num == 1:
            page_url = base_url
        else:
            page_url = f"{base_url}p_{page_num}/"

        response = requests.get(page_url)
        soup = BeautifulSoup(response.content, 'html.parser')

        agent_divs = soup.select('.split-4.agent-team-results .name.notranslate')

        for agent_div in agent_divs:
            agent_a = agent_div.find('a')
            if agent_a and 'href' in agent_a.attrs:
                agent_url = f"https://www.coldwellbankerhomes.com{agent_a['href']}"
                agent_data = scrape_page(agent_url)
                all_agent_data.append(agent_data)

    return all_agent_data

# Web address to scrape
web_address = "https://www.coldwellbankerhomes.com/sc/columbia/agents/"
start_page = 1  # Specify the start page number
end_page = 25    # Specify the end page number

agent_data_list = scrape_specific_pages(web_address, start_page, end_page)

# Declare the output file path
output_file_path = r'D:\Office\Scripts\output_1.xlsx'

# Check if the output file already exists
if os.path.exists(output_file_path):
    print("Output file already exists. Please remove or rename the existing file.")
else:
    wb = openpyxl.Workbook()
    ws = wb.active

    header = ['Name', 'Role', 'Email', 'Mobile', 'Office', 'Direct', 'Market', 'Company', 'Address']
    ws.append(header)

    for agent_data in agent_data_list:
        row_data = [agent_data.get(key, '') for key in header]
        ws.append(row_data)

    wb.save(output_file_path)
    print(f"Data saved to {output_file_path}")
    print(f"Date Scraped {len(agent_data_list)}")


# In[ ]:




