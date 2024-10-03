##########################################################################################################
##########################################################################################################
import requests
import pandas as pd
from datetime import datetime

# Load the Excel sheet with faculty names and Scopus IDs
faculty_data = pd.read_excel('faculty_scopus_id_list.xlsx')  # Adjust file name if necessary

# Define your API key
########################################
API_KEY =  'd72ef41f58ca547602b39d5be185b011'
########################################
def get_dates():
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
            end_year = int(input("Please enter the end year (e.g., 2024): "))
            break
        except ValueError:
            print("Invalid input. Please enter an integer.")

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
            start_year = int(input("Please enter the start year (e.g., 2024): "))
            if start_year <= end_year:
                break
            else:
                print(f"The start year must be less than or equal to the end year ({end_year}).")
        except ValueError:
            print("Invalid input. Please enter an integer.")

    # Adjust years for API query if necessary
    query_start_year = start_year
    query_end_year = end_year
    if start_year == end_year:
        query_start_year = start_year - 1
        #print(f"Note: For API query, start year adjusted to {query_start_year} to ensure different years.")

    return start_month, start_year, end_month, end_year, query_start_year, query_end_year

# Get user input for dates
start_month, start_year, end_month, end_year, query_start_year, query_end_year = get_dates()
print(f"You entered: Start Date: {start_month}/{start_year}, End Date: {end_month}/{end_year}")
#print(f"API query will use: Start Year: {query_start_year}, End Year: {query_end_year}")
print("\nScraping the data from Scopus\n")
print("NOTE: To see the list of faculties whose data is being scraped, open the adjacent Excel file named 'faculty_scopus_id_list'. You can add or remove names in the future from this Excel file.")
print("\nWorking on scraping...")

all_papers = []

# Looping through each faculty member and their Scopus ID
for index, row in faculty_data.iterrows():
    faculty_name = row['faculty']
    scopus_id = row['scopus']
    # API request URL for each faculty member
    base_url = f'https://api.elsevier.com/content/search/scopus?apiKey={API_KEY}&query=AU-ID({scopus_id})&date={query_start_year}-{end_year}'
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
                    year = publication_date.year
                    month = publication_date.month

                    # Filtering only if publication falls within the original input months and year
                    #if (year == end_year and start_month <= month <= end_month):
                    if (start_year <= year <= end_year and start_month <= month <= end_month):
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

df_all_papers.to_excel(f'faculty_scraped_data_{start_month}_{start_year}_to_{end_month}_{end_year}.xlsx', index=False)
print("\nData saved to Excel file")
###############################################YashUpase##################################################
##########################################################################################################
