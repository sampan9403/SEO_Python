from serpapi import GoogleSearch
import pandas as pd
from datetime import datetime
import concurrent.futures
from tqdm import tqdm
from urllib.parse import unquote
import os

def fetch_data(keyword):
    all_results = []
    current_position = 1  # Initialize a counter for position

    for offset in range(0, 100, 20):  # Iterate over pages (0, 20, 40, 60, 80)
        params = {
            "engine": "google_maps",
            "q": keyword,
            "type": "search",
            "ll": "@22.3193039,114.1693611,12z",
            "hl": "zh-tw",
            "google_domain": "google.com.hk",
            "start": offset,  # Set the offset for pagination
            "api_key": "7ec738d5269f13638e91542f7419b4122c614ac0e82fcc4ee5afbfd072f910cb"
        }

        search = GoogleSearch(params)
        results = search.get_dict()
        local_results = results.get("local_results", [])

        if not local_results:
            break  # Break the loop if no results are returned

        for result in local_results:
            title = unquote(result.get("title", ""))
            all_results.append({"Keyword": keyword, "Position": current_position, "Title": title})
            current_position += 1  # Increment the position

            if len(all_results) >= 100:  # Stop if 100 results are accumulated
                break

        if len(all_results) >= 100:
            break

    return all_results

# Function to filter SERP data based on titles and add unmatched or no-result titles
def process_data(serp_data, keywords, titles):
    matched_data = set()
    filtered_data = []

    for item in serp_data:
        for title in titles:
            if title.lower() in item['Title'].lower():
                matched_data.add(item['Keyword'])
                filtered_data.append(item)
                break

    # Add unmatched keywords with NA values or no-result keywords with specific message
    for keyword in keywords:
        if keyword not in matched_data:
            if any(d['Keyword'] == keyword for d in serp_data):
                filtered_data.append({"Keyword": keyword, "Position": "NA", "Title": "NA"})
            else:
                filtered_data.append({"Keyword": keyword, "Position": "沒有相關搜索結果", "Title": "沒有相關搜索結果"})

    return filtered_data

# Main function
def main():
    # Load keywords from Excel
    df_keywords = pd.read_excel("keyword_list_google_map_serp.xlsx", sheet_name="keywords")
    keywords = df_keywords["keywords"].tolist()

    # Load titles from Excel
    df_titles = pd.read_excel("keyword_list_google_map_serp.xlsx", sheet_name="title")
    titles = df_titles["title"].tolist()

    # Prepare for concurrent fetching
    all_data = []
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [executor.submit(fetch_data, keyword) for keyword in keywords]
        for future in tqdm(concurrent.futures.as_completed(futures), total=len(keywords)):
            all_data.extend(future.result())

    # Create DataFrame from fetched data
    df_results = pd.DataFrame(all_data)

    # Process data for ranking worksheet
    processed_data = process_data(all_data, keywords, titles)
    df_processed = pd.DataFrame(processed_data)

    # Export to Excel with timestamp and custom sheet names
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"google_maps_results_{timestamp}.xlsx"
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        df_results.to_excel(writer, sheet_name='SERP', index=False)
        df_processed.to_excel(writer, sheet_name='ranking', index=False)
    print(f"Data exported to {filename}")

if __name__ == "__main__":
    main()
