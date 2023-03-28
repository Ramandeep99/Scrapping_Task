import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

# Load the Google Sheet into a pandas DataFrame
# print(os.path.join('/'))
df = pd.read_excel(os.path.join('./Web Scraping Assignment.xlsx'))

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    # Extract the website URL for the corresponding college
    if(pd.isnull(row["WEBSITE"])):
        continue
    website = row['WEBSITE']
    
    # Fetch the HTML content of the website
    response = requests.get(website)
    html_content = response.text
    
    # # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # # Find the HTML element that contains the address information of the college
    address_element = soup.find('address')
    
    # # Extract the text content of the address element
    if address_element:
        address = address_element.text.strip()
    else:
        address = ''
    
    # Write the address to the corresponding cell in the DataFrame
    df.loc[index, 'Address'] = address
    # print(address)

# Save the updated DataFrame to an Excel file
df.to_excel('colleges_with_addresses.xlsx', index=False)
