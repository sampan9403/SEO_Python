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
            link_chinese = unquote(link)
            
            # Modify the domain check
            if f"://{domain}" in link or f".{domain}" in link or \
                f"://{domain}" in link_chinese or f".{domain}" in link_chinese:
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
domains_list = keywords_df["domains"].apply(lambda x: [d.strip() for d in x.split(',')])

# SERPAPI parameters common to all searches
base_params = {
    "location": "Hong Kong",
    "google_domain": "google.com",
    "gl": "hk",
    "api_key": "7ec738d5269f13638e91542f7419b4122c614ac0e82fcc4ee5afbfd072f910cb",
    "num": 100  # Adjust this number as needed
}

# Initialize DataFrame for storing rankings
rankings_df = pd.DataFrame(columns=["Query", "Domain", "Position", "Title", "Link", "Link(Chinese)", "Display Link", "Snippet"])

# Use ThreadPoolExecutor to fetch data concurrently
results_dict = {}
with concurrent.futures.ThreadPoolExecutor() as executor:
    future_to_query_domain = {}
    for query, domains in zip(keywords, domains_list):
        for domain in domains:
            future = executor.submit(fetch_serp_data, query, domain)
            future_to_query_domain[future] = (query, domain)

    for future in tqdm(concurrent.futures.as_completed(future_to_query_domain), total=sum(len(domains) for domains in domains_list), desc="Fetching SERP Data"):
        query, domain = future_to_query_domain[future]
        result = future.result()
        results_dict[(query, domain)] = result

# Append results to DataFrame in the order of the input file
for query, domains in zip(keywords, domains_list):
    for domain in domains:
        rankings_df = pd.concat([rankings_df, results_dict[(query, domain)]], ignore_index=True)

# Create a list of unique domains in the order they first appear
unique_domains_ordered = []
for domain_list in domains_list:
    for domain in domain_list:
        domain_clean = domain.split('://')[-1]
        if domain_clean not in unique_domains_ordered:
            unique_domains_ordered.append(domain_clean)

# Prepare the 'Ranking_Only' table with ordered columns and no duplicate keywords
ranking_only_df = pd.DataFrame(index=pd.unique(keywords), columns=unique_domains_ordered)
for query, domains in zip(keywords, domains_list):
    for domain in domains:
        domain_clean = domain.split('://')[-1]
        positions = rankings_df[(rankings_df['Query'] == query) & (rankings_df['Domain'] == domain)]['Position'].tolist()
        positions = [str(pos) for pos in positions if pos != 'no result']
        ranking_only_df.at[query, domain_clean] = ','.join(positions) if positions else 'NA'

# Get current time
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# Export the data to Excel
with pd.ExcelWriter(f"keyword_search_rankings_{current_time}.xlsx", engine='xlsxwriter') as writer:
    # Export 'Ranking_SERP' Worksheet
    rankings_df.to_excel(writer, index=False, sheet_name='Ranking_SERP')
    
    # Export 'Ranking_Only' Worksheet
    ranking_only_df.to_excel(writer, startrow=1, startcol=1, sheet_name='Ranking_Only')

    # Access the workbook and the worksheet objects
    workbook  = writer.book
    worksheet = writer.sheets['Ranking_Only']

    # Define a format for bold text and center alignment
    bold_center_format = workbook.add_format({'bold': True, 'align': 'center'})

    # Write "keywords\domains" in the top-left cell (B2)
    worksheet.write('B2', 'keywords\\domains', bold_center_format)

    # Adjust the column widths
    for idx, col in enumerate(ranking_only_df.columns):
        # Determine the maximum width of the column based on the cell contents
        column_len = max(ranking_only_df[col].astype(str).map(len).max(), len(str(col))) + 2
        worksheet.set_column(idx + 1, idx + 1, column_len)  # Set column width
