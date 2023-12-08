from serpapi import GoogleSearch
import pandas as pd
from datetime import datetime
import concurrent.futures
from tqdm import tqdm
from urllib.parse import unquote

# Function to fetch SERP data for a single query and domain
def fetch_serp_data(query, domain):
    search_params = base_params.copy()
    search_params["q"] = query

    try:
        search_results = GoogleSearch(search_params).get_dict()
        organic_results = search_results.get("organic_results", [])
    except Exception as e:
        print(f"Error fetching results for query '{query}': {e}")
        organic_results = []  # Ensure empty results in case of an error

    temp_df = pd.DataFrame(columns=["Query", "Domain", "Position", "Title", "Link", "Link(Chinese)", "Display Link", "Snippet"])

    if not organic_results:  # Check if organic_results is empty
        # Insert a placeholder row for 'no result'
        temp_df = pd.concat([temp_df, pd.DataFrame({
            "Query": [query],
            "Domain": [domain],
            "Position": ["no result"],
            "Title": ["no result"],
            "Link": ["no result"],
            "Link(Chinese)": ["no result"],
            "Display Link": ["no result"],
            "Snippet": ["no result"]
        })], ignore_index=True)
    else:
        found_domain_result = False
        for i, result in enumerate(organic_results):
            link = result.get("link", "")
            link_chinese = unquote(link)  # Decode percent-encoded URL to Chinese characters

            if domain in link or domain in link_chinese:
                found_domain_result = True
                position = i + 1
                title = result.get("title", "")
                display_link = result.get("displayed_link", "")
                snippet = result.get("snippet", "")

                temp_df = pd.concat([temp_df, pd.DataFrame({
                    "Query": [query],
                    "Domain": [domain],
                    "Position": [position],
                    "Title": [title],
                    "Link": [link],
                    "Link(Chinese)": [link_chinese],
                    "Display Link": [display_link],
                    "Snippet": [snippet]
                })], ignore_index=True)

        # Check if no results were found for the specified domain
        if not found_domain_result:
            temp_df = pd.concat([temp_df, pd.DataFrame({
                "Query": [query],
                "Domain": [domain],
                "Position": ["no result"],
                "Title": ["no result"],
                "Link": ["no result"],
                "Link(Chinese)": ["no result"],
                "Display Link": ["no result"],
                "Snippet": ["no result"]
            })], ignore_index=True)

    return temp_df

# Read keywords and domains from an Excel file
keywords_df = pd.read_excel("keyword_ranking_list.xlsx")
keywords = keywords_df["keywords"].tolist()
domains_list = keywords_df["domains"].apply(lambda x: x.split(','))  # Assuming domains are comma-separated

# SERPAPI parameters common to all searches
base_params = {
    "location": "Hong Kong",
    "google_domain": "google.com.hk",
    "hl": "en",
    "gl": "hk",
    "api_key": "7ec738d5269f13638e91542f7419b4122c614ac0e82fcc4ee5afbfd072f910cb",
    "num": 100  # Adjust this number as needed
}

# Initialize DataFrame for storing rankings
rankings_df = pd.DataFrame(columns=["Query", "Domain", "Position", "Title", "Link", "Link(Chinese)", "Display Link", "Snippet"])

# Use ThreadPoolExecutor to fetch data concurrently
with concurrent.futures.ThreadPoolExecutor() as executor:
    future_to_query_domain = {}
    for query, domains in zip(keywords, domains_list):
        for domain in domains:
            future = executor.submit(fetch_serp_data, query, domain.strip())  # Strip to remove any leading/trailing whitespace
            future_to_query_domain[future] = (query, domain)

    for future in tqdm(concurrent.futures.as_completed(future_to_query_domain), total=sum(len(domains) for domains in domains_list), desc="Fetching SERP Data"):
        query, domain = future_to_query_domain[future]
        result = future.result()
        rankings_df = pd.concat([rankings_df, result], ignore_index=True)

# Get current time
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# Export the rankings DataFrame to an Excel file with the current time in the filename
with pd.ExcelWriter(f"keyword_search_rankings_{current_time}.xlsx") as writer:
    rankings_df.to_excel(writer, index=False)
