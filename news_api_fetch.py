# pulls the latest business top news from newsapi, parses for company name, stores it into a excel file
# can be scheduled for incremental updates to local excel file
# TODO: add keyword filtering (e.g. layoffs, founding round, etc)

import requests
import pandas as pd
import re
from datetime import datetime
import schedule
import time

def get_business_news(api_key, language='en', country='us'):
    url = 'https://newsapi.org/v2/top-headlines'
    params = {
        'apiKey': api_key,
        'category': 'business',
        'language': language,
        'country': country,
        'pageSize': 70  
    }

    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()

        if data['status'] == 'ok':
            return data['articles']
        else:
            print('Error: Unable to fetch news.')
            return []
    except requests.exceptions.RequestException as e:
        print(f'Error: {e}')
        return []

def extract_company_names(article_text):
    # find all uppercase words longer than two characters.
    potential_company_names = re.findall(r'\b[A-Z]{3,}\b', article_text)
    return potential_company_names

def create_dataframe(articles):
    data = []

    for article in articles:
        title = article['title']
        description = article['description'] if article['description'] else ''
        article_text = title + ' ' + description
        potential_company_names = extract_company_names(article_text)[:5]  # maximum five companies.

        summary = f"{title} - {description}" if description else title

        row_data = {
            'Title': title,
            'Summary': summary,
            'URL': article['url'],
            'Published At': article['publishedAt'],  # raw UTC time, removing timezone 
        }

        # five company columns, filling with None if fewer than five.
        for i in range(5):
            if i < len(potential_company_names):
                row_data[f'Mentioned Company {i + 1}'] = potential_company_names[i]
            else:
                row_data[f'Mentioned Company {i + 1}'] = None

        data.append(row_data)

    if data:
        df = pd.DataFrame(data)
        return df
    else:
        return None

def update_excel_file(df):
    try:
        df_existing = pd.read_excel('business_news2.xlsx')

        # new stories by comparing the 'Title' column with the existing DataFrame.
        new_stories = df[~df['Title'].isin(df_existing['Title'])]

        if not new_stories.empty:
            df_existing = pd.concat([df_existing, new_stories], ignore_index=True, sort=False)
            df_existing.to_excel('business_news.xlsx', index=False)
            print("New stories added to Excel file.")
        else:
            print("No new stories found.")

    except FileNotFoundError:
        df.to_excel('business_news.xlsx', index=False)
        print("New Excel file created with the current data.")

def fetch_and_update():
    api_key = ''  #API key.

    print("Fetching business news...\n")
    articles = get_business_news(api_key)
    if articles:
        print("Top business news headlines:\n")
        df = create_dataframe(articles)
        print(df)
        update_excel_file(df)
    else:
        print("No news articles found.")

if __name__ == '__main__':
    # Schedule to run every xx minutes.
    schedule.every(1).minutes.do(fetch_and_update)

    while True:
        schedule.run_pending()
        time.sleep(1)
