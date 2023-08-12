import pandas as pd
from bs4 import BeautifulSoup

def extract_titles_from_url(file_name):
    """Extract titles (from URLs) and URLs from the given HTML file."""
    with open(file_name, 'r', encoding='utf-8') as file:
        content = file.read()
        soup = BeautifulSoup(content, 'html.parser')
        
        results = []
        for a_tag in soup.find_all('a', href=True):
            # Filter out unwanted links
            if 'http' in a_tag['href'] and 'google.com' not in a_tag['href']:
                # Extract title from URL after the "›"
                title = a_tag.text.split('›')[-1].strip()
                url = a_tag['href']
                results.append((title, url))
    return results

def create_hyperlink_spreadsheet_from_url(file_names, output_file):
    """Create an Excel spreadsheet with hyperlinks based on the provided HTML files (using URL for title)."""
    all_results = []
    for file_name in file_names:
        all_results.extend(extract_titles_from_url(file_name))
    
    df_hyperlinks = pd.DataFrame({
        'Hyperlinks': ['=HYPERLINK("{}", "{}")'.format(url, title) for title, url in all_results]
    })
    
    df_hyperlinks.to_excel(output_file, index=False, engine='openpyxl')

if __name__ == '__main__':
    # List of all HTML file names (adjust this to match the number of files you have)
    file_names = [f"./{i}.html" for i in range(1, 11)]
    
    # Create the Excel spreadsheet using URL for title
    output_file = "./search_results_url_hyperlinks.xlsx"
    create_hyperlink_spreadsheet_from_url(file_names, output_file)

