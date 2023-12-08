from serpapi import GoogleSearch
import pandas as pd
from datetime import datetime
import urllib.parse
from tqdm import tqdm

# Read keywords from the Excel file
keywords_df = pd.read_excel('keyword_list.xlsx')
keywords = keywords_df['keywords'].tolist()

# Initialize DataFrame to store the rankings
rankings_df = pd.DataFrame(columns=["Query", "Related Query", "Position", "Title", "Link", "Link(Chinese)", "Display Link", "Snippet"])

for keyword in tqdm(keywords, desc="Processing Keywords"):
    params = {
        "q": keyword,
        "location": "Hong Kong",
        "google_domain": "google.com.hk",
        "hl": "en",
        "gl": "hk",
        "api_key": "e544226a57b6e485eeae6b872df1b64194d27f9dad5eedfb0f2cb0022f1fcefa",
        "num": 100
    }

    # Execute SERPAPI request for the main query
    search = GoogleSearch(params)
    main_results = search.get_dict().get("organic_results", [])

    # Process main query results
    for i, result in enumerate(main_results):
        new_row = pd.DataFrame({
            "Query": [keyword],
            "Related Query": [keyword],
            "Position": [i + 1],
            "Title": [result.get("title", "")],
            "Link": [result.get("link", "")],
            "Link(Chinese)": [urllib.parse.unquote(result.get("link", ""))],
            "Display Link": [result.get("displayed_link", "")],
            "Snippet": [result.get("snippet", "")]
        })
        rankings_df = pd.concat([rankings_df, new_row], ignore_index=True)

    # Process related searches
    related_searches = search.get_dict().get("related_searches", [])
    for related_search in related_searches:
        related_query = related_search['query']
        related_search_params = params.copy()
        related_search_params['q'] = related_query

        related_search_results = GoogleSearch(related_search_params).get_dict().get("organic_results", [])
        for j, rel_result in enumerate(related_search_results):
            new_row = pd.DataFrame({
                "Query": [keyword],
                "Related Query": [related_query],
                "Position": [j + 1],
                "Title": [rel_result.get("title", "")],
                "Link": [rel_result.get("link", "")],
                "Link(Chinese)": [urllib.parse.unquote(rel_result.get("link", ""))],
                "Display Link": [rel_result.get("displayed_link", "")],
                "Snippet": [rel_result.get("snippet", "")]
            })
            rankings_df = pd.concat([rankings_df, new_row], ignore_index=True)

# Get current time
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# Export the rankings DataFrame to an Excel file with the current time in the filename
output_filename = f"related_search_rankings_{current_time}.xlsx"
with pd.ExcelWriter(output_filename) as writer:
    rankings_df.to_excel(writer, index=False)
