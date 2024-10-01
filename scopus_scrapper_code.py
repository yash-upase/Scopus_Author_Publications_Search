import requests
import pandas as pd
from datetime import datetime

# Load the Excel sheet with faculty names and Scopus IDs
faculty_data = pd.read_excel('faculty_scopus_id_list.xlsx')  # Adjust file name if necessary

# Define your API key
API_KEY = 'add_your_api_key_here'



def get_dates():
    while True:
        try:
            start_month = int(input("Please enter the start month (1-12): "))
            if 1 <= start_month <= 12:
                break
            else:
                print("Please enter a valid month (1-12).")
        except ValueError:
            print("Invalid input. Please enter an integer.")

    while True:
        try:
            start_year = int(input("Please enter the start year (e.g., 2023): "))
            break  # No specific range check for year; can be any integer
        except ValueError:
            print("Invalid input. Please enter an integer.")

    while True:
        try:
            end_month = int(input("Please enter the end month (1-12): "))
            if 1 <= end_month <= 12:
                break
            else:
                print("Please enter a valid month (1-12).")
        except ValueError:
            print("Invalid input. Please enter an integer.")

    while True:
        try:
            end_year = int(input("Please enter the end year (e.g., 2023): "))
            # Ensure the end year is greater than or equal to the start year
            if end_year > start_year or (end_year == start_year and end_month >= start_month):
                break
            else:
                print(f"The end year must be greater than or equal to the start year ({start_year}), "
                      f"or if they are the same, the end month must be greater than or equal to the start month ({start_month}).")
        except ValueError:
            print("Invalid input. Please enter an integer.")

    return start_month, start_year, end_month, end_year

start_month, start_year, end_month, end_year = get_dates()
print(f"You entered: Start Date: {start_month}/{start_year}, End Date: {end_month}/{end_year}")
print("Scrapping the data from scopus")

all_papers = []

# Looping through each faculty member and their Scopus ID
for index, row in faculty_data.iterrows():
    faculty_name = row['faculty']
    scopus_id = row['scopus']
    # API request URL for each faculty member
    base_url = f'https://api.elsevier.com/content/search/scopus?apiKey={API_KEY}&query=AU-ID({scopus_id})&date={start_year}-{end_year}'
    
    # Make the request
    response = requests.get(base_url)
    
    # Check response status
    if response.status_code == 200:
        data = response.json()

        # Extracting paper details
        for entry in data.get('search-results', {}).get('entry', []):
            publication_date_str = entry.get('prism:coverDate', 'N/A')
            if publication_date_str != 'N/A':
                try:
                    
                    publication_date = datetime.strptime(publication_date_str, "%Y-%m-%d")
                    month = publication_date.month

                    # Filter by month 
                    if start_month <= month <= end_month:
                        paper_info = {
                            'Faculty': faculty_name,
                            'Title': entry.get('dc:title', 'N/A'),
                            'Authors': entry.get('dc:creator', 'N/A'),
                            'Publication Date': publication_date_str,
                            'Source Title': entry.get('prism:publicationName', 'N/A'),
                            'DOI': entry.get('prism:doi', 'N/A')
                            #'Abstract': entry.get('dc:description', 'N/A')
                        }
                        all_papers.append(paper_info)
                except ValueError:
                    print(f"Skipping entry for {faculty_name} due to date parsing error: {publication_date_str}")
    else:
        print(f"Error for {faculty_name}: {response.status_code}, {response.text}")

df_all_papers = pd.DataFrame(all_papers)

df_all_papers.to_excel(f'scopus_papers_all_faculties_{start_month}_{start_year}_to_{end_month}_{end_year}.xlsx', index=False)
print("Data saved to excel file")
#YashUpase_Jay
