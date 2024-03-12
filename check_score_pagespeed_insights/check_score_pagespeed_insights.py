from googleapiclient.discovery import build
import pandas as pd
from datetime import datetime
from tqdm import tqdm
import concurrent.futures


# Your API key
API_KEY = 'AIzaSyDkO_4lCsanUQp7ihIaIMxrbpZkH79Yb50'

# Initialize the service
service = build('pagespeedonline', 'v5', developerKey=API_KEY)

# Read URLs from Excel
df_urls = pd.read_excel('url_list_pagespeed.xlsx', sheet_name='url', usecols=['url'])

def format_issue(issue):
    return f"{issue['title']}\n{issue['description']}\nEstimatedSavings: {issue['estimatedSavings']} ms"

def analyze_url(url, strategy):
    result = service.pagespeedapi().runpagespeed(url=url, strategy=strategy).execute()
    lighthouse_result = result.get('lighthouseResult', {})
    audits = lighthouse_result.get('audits', {})
    categories = lighthouse_result.get('categories', {})
    performance_score = categories.get('performance', {}).get('score') * 100

    metrics = {key: audits[key]['displayValue'] for key in ['first-contentful-paint', 'largest-contentful-paint', 'total-blocking-time', 'cumulative-layout-shift', 'speed-index']}

    opportunities = {}
    for opp_id, opp in audits.items():
        if opp.get('score') is None or 'details' not in opp or 'overallSavingsMs' not in opp['details']:
            continue
        opportunities[opp_id] = {
            'title': opp['title'],
            'description': opp.get('description', ''),
            'estimatedSavings': opp['details'].get('overallSavingsMs', 0)
        }
    top_issues = sorted(opportunities.items(), key=lambda item: item[1]['estimatedSavings'], reverse=True)[:3]
    formatted_issues = [format_issue(issue[1]) for issue in top_issues]

    return {
        'url': url,
        'strategy': strategy,
        'score_of_performance': performance_score,
        **metrics,
        'issue 1': formatted_issues[0] if len(formatted_issues) > 0 else "",
        'issue 2': formatted_issues[1] if len(formatted_issues) > 1 else "",
        'issue 3': formatted_issues[2] if len(formatted_issues) > 2 else "",
        'check_date': datetime.now().strftime('%Y-%m-%d')
    }

def main():
    results = []
    with concurrent.futures.ThreadPoolExecutor() as executor:
        future_to_url = {executor.submit(analyze_url, url[0], strategy): url for url in df_urls.itertuples(index=False) for strategy in ['MOBILE', 'DESKTOP']}
        for future in tqdm(concurrent.futures.as_completed(future_to_url), total=len(future_to_url), desc="Analyzing URLs"):
            results.append(future.result())

    results_df = pd.DataFrame(results)
    results_df.sort_values(by=['url'], inplace=True)

    # Exporting results
    timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
    results_df.to_excel(f'pagespeed_insights_performance-{timestamp}.xlsx', index=False, sheet_name='pagespeed-insights')

if __name__ == "__main__":
    main()
