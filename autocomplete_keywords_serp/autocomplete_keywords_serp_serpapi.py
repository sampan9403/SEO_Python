from serpapi import GoogleSearch
import pandas as pd
from datetime import datetime
import urllib.parse
from tqdm import tqdm

# Read keywords from the Excel file
keywords_df = pd.read_excel('keyword_list_autocomplete.xlsx')
keywords = keywords_df['keywords'].tolist()

# Initialize DataFrame to store the rankings
rankings_df = pd.DataFrame(columns=["Query", "Autocomplete Query", "Position", "Title", "Link", "Link(Chinese)", "Display Link", "Snippet"])

for keyword in tqdm(keywords, desc="Processing Keywords"):
    params = {
        "q": keyword,
        "location": "Hong Kong",
        "google_domain": "google.com",
        "gl": "hk",
        "api_key": "7ec738d5269f13638e91542f7419b4122c614ac0e82fcc4ee5afbfd072f910cb", # info@owlish
        "num": 100
    }

    # Process the original query
    original_search = GoogleSearch(params)
    original_results = original_search.get_dict().get("organic_results", [])

    for i, result in enumerate(original_results):
        new_row = pd.DataFrame({
            "Query": [keyword],
            "Autocomplete Query": [keyword], # Original query in the Autocomplete Query column
            "Position": [i + 1],
            "Title": [result.get("title", "")],
            "Link": [result.get("link", "")],
            "Link(Chinese)": [urllib.parse.unquote(result.get("link", ""))],
            "Display Link": [result.get("displayed_link", "")],
            "Snippet": [result.get("snippet", "")]
        })
        rankings_df = pd.concat([rankings_df, new_row], ignore_index=True)

    # Fetch Google Autocomplete suggestions
    autocomplete_search = GoogleSearch({**params, "engine": "google_autocomplete"})
    autocomplete_results = autocomplete_search.get_dict().get("suggestions", [])

    for autocomplete_result in autocomplete_results:
        autocomplete_query = autocomplete_result['value']

        # Execute SERPAPI request for each autocomplete suggestion
        search_params = {**params, "q": autocomplete_query}
        search_results = GoogleSearch(search_params).get_dict().get("organic_results", [])

        for i, result in enumerate(search_results):
            new_row = pd.DataFrame({
                "Query": [keyword],
                "Autocomplete Query": [autocomplete_query],
                "Position": [i + 1],
                "Title": [result.get("title", "")],
                "Link": [result.get("link", "")],
                "Link(Chinese)": [urllib.parse.unquote(result.get("link", ""))],
                "Display Link": [result.get("displayed_link", "")],
                "Snippet": [result.get("snippet", "")]
            })
            rankings_df = pd.concat([rankings_df, new_row], ignore_index=True)

# Get current time
current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# Export the rankings DataFrame to an Excel file with the current time in the filename
output_filename = f"autocomplete_search_rankings_{current_time}.xlsx"
with pd.ExcelWriter(output_filename) as writer:
    rankings_df.to_excel(writer, index=False)
