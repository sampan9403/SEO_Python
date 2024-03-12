import requests
from xml.etree import ElementTree
import pandas as pd
from datetime import datetime
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor

# Define a threshold for the maximum number of URLs per sheet
MAX_URLS_PER_SHEET = 30000

# Function to fetch and parse a sitemap, returning URL and lastmod
def fetch_sitemap_details(sitemap_url):
    try:
        response = requests.get(sitemap_url)
        root = ElementTree.fromstring(response.content)
        namespace = {'ns': 'http://www.sitemaps.org/schemas/sitemap/0.9'}
        urls = [
            (
                url.find('ns:loc', namespace).text, 
                url.find('ns:lastmod', namespace).text if url.find('ns:lastmod', namespace) is not None else "None", 
                sitemap_url
            ) 
            for url in root.findall('ns:url', namespace)
        ]
        return urls
    except Exception as e:
        print(f"Error fetching or parsing {sitemap_url}: {e}")
        return []

# Adjusted process_sitemaps to update progress per URL
def process_sitemaps(sitemap_urls):
    all_urls = []
    for sitemap_url in tqdm(sitemap_urls, desc="Processing Sitemaps"):
        urls = fetch_sitemap_details(sitemap_url)
        all_urls.extend(urls)
    return all_urls

# Function to split URLs into multiple sheets if exceeding MAX_URLS_PER_SHEET
def export_to_excel(url_details):
    timestamp = datetime.now().strftime('%Y-%m-%d_%H%M%S')
    filename = f'ExportedURLs_{timestamp}.xlsx'
    writer = pd.ExcelWriter(filename, engine='openpyxl')
    
    # Determine the number of sheets needed
    total_urls = len(url_details)
    num_sheets = total_urls // MAX_URLS_PER_SHEET + (1 if total_urls % MAX_URLS_PER_SHEET > 0 else 0)
    
    for sheet_number in range(1, num_sheets + 1):
        start_index = (sheet_number - 1) * MAX_URLS_PER_SHEET
        end_index = min(sheet_number * MAX_URLS_PER_SHEET, total_urls)
        df_sheet = pd.DataFrame(url_details[start_index:end_index], columns=['URL', 'Last Modified', 'Sitemap'])
        sheet_name = f'{str(sheet_number).zfill(2)}'  # Sheet names as '01', '02', ...
        df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)
    
    writer.save()
    print(f"Exported to {filename}")

# Main function to orchestrate the workflow
def main():
    # Read sitemap URLs from Excel
    df_sitemaps = pd.read_excel('sitemaps.xlsx', sheet_name='sitemaps')
    sitemap_urls = df_sitemaps['sitemaps'].tolist()

    # Process sitemaps and show detailed progress
    url_details = process_sitemaps(sitemap_urls)

    # Export to Excel, splitting into multiple sheets if necessary
    export_to_excel(url_details)

if __name__ == "__main__":
    main()
