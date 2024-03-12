import pandas as pd
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
from tqdm import tqdm  # tqdm is used to show a smart progress meter
from urllib.parse import unquote

headers = {
    'User-Agent': 'Mozilla/5.0',
    'Accept': '*/*',
}

# Change the file name here
path = 'url_canonical.xlsx'
timeout = 10  # Adjust the timeout as needed
max_workers = 10  # Adjust the number of workers as needed

def requestUrl(url):
    try:
        response = requests.get(url, headers=headers, allow_redirects=True, timeout=timeout)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            canonical_tag = soup.find('link', rel='canonical')
            canonical_url = canonical_tag['href'] if canonical_tag else ''

            # Decode URLs to ensure they are compared in the same format
            decoded_response_url = unquote(response.url)
            decoded_canonical_url = unquote(canonical_url)

            # Compare the decoded URLs
            comparison = "Same" if (decoded_canonical_url and decoded_response_url == decoded_canonical_url) else "Not Same"
            return decoded_response_url, response.status_code, decoded_canonical_url, comparison
        else:
            return unquote(response.url), response.status_code, 'Error or no response', 'Not Applicable'
    except requests.exceptions.Timeout:
        return "Timeout", "Timeout", "Timeout", "Not Applicable"

# Read data from the specified Excel file
data = pd.read_excel(path)

# Prepare your URLs list from the dataframe
urls = data['original url'].tolist()

results = []

# Setup progress bar with 'tqdm'
with ThreadPoolExecutor(max_workers=max_workers) as executor:
    # Prepare the futures dictionary
    future_to_url = {executor.submit(requestUrl, url): url for url in urls}
    # Iterate through the futures as they complete (as_completed)
    for future in tqdm(as_completed(future_to_url), total=len(urls), desc="Checking URLs"):
        url = future_to_url[future]
        try:
            redirect_url, status_code, canonical_url, comparison = future.result()
            results.append({
                "original url": url,
                "redirect to": redirect_url,
                "status code": status_code,
                "canonical url": canonical_url,
                "Same for original and canonical?": comparison
            })
        except Exception as e:
            results.append({
                "original url": url,
                "redirect to": "Error",
                "status code": "Error",
                "canonical url": "Error",
                "Same for original and canonical?": "Not Applicable"
            })

result_df = pd.DataFrame(results)

time_str = time.strftime("%Y-%m-%d_%H_%M_%S", time.localtime())
save_file_name = f'result_canonical_{time_str}.xlsx'
result_df.to_excel(save_file_name, index=False)

print(f'Result Saved To: {save_file_name}')
