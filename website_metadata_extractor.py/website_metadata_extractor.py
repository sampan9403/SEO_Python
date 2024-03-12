import pandas as pd
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm
from datetime import datetime

# Load URLs from an Excel file
excel_name = 'url_check_HTML.xlsx'
df_urls = pd.read_excel(excel_name, sheet_name='url')

# Initialize a requests session for efficiency
session = requests.Session()

# Prepare an empty list to store the results
results = []

# Process each URL in the DataFrame
for url in tqdm(df_urls['url'], desc="Checking URLs"):
    # Initialize result dictionary with URL
    result = {'URL': url}
    
    try:
        # Sending a request to the URL
        response = session.get(url)
        
        # Creating a BeautifulSoup object and specifying the parser
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Extracting the title tag
        result['Title'] = soup.title.string.replace('\n', ' ').strip() if soup.title else 'NAN'
        
        # Extracting the meta description
        meta_description = soup.find('meta', attrs={'name': 'description'})
        result['Meta Description'] = meta_description['content'].replace('\n', ' ').strip() if meta_description else 'NAN'
        
        # Extracting H1 and H2 tags
        result['H1'] = '\n'.join([h1.get_text().replace('\n', ' ').strip() for h1 in soup.find_all('h1')]) or 'NAN'
        result['H2'] = '\n'.join([h2.get_text().replace('\n', ' ').strip() for h2 in soup.find_all('h2')]) or 'NAN'
        
    except Exception as e:
        result.update({'Title': 'NAN', 'Meta Description': 'NAN', 'H1': 'NAN', 'H2': 'NAN'})
    
    results.append(result)

# Convert the results into a DataFrame
df_results = pd.DataFrame(results)

# Export the results to an Excel file
current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
output_excel_name = f'url_check_results_{current_time}.xlsx'
df_results.to_excel(output_excel_name, index=False)

print(f'Results have been saved to {output_excel_name}')
