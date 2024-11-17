import pandas as pd
from bs4 import BeautifulSoup

# File paths
html_file = 'metacomplexity.html'
excel_file = 'papers.xlsx'

# Read the Excel file
df = pd.read_excel(excel_file)

# Verify that the necessary columns exist
required_columns = {'Year', 'Title', 'Authors', 'Publication', 'Link', 'Summary', 'Techniques'}
if not required_columns.issubset(df.columns):
    raise ValueError(f"The Excel file must contain the following columns: {required_columns}")

# Load the HTML content
with open(html_file, 'r', encoding='utf-8') as file:
    soup = BeautifulSoup(file, 'html.parser')

# Locate the bibliography section or fallback to the container
container = soup.find('div', class_='container')
bibliography_header = container.find('h3', string='Bibliography')
bibliography_section = bibliography_header.find_next_sibling() if bibliography_header else None

if not bibliography_section:
    # Default to appending directly to the container if no sibling exists
    bibliography_section = container

# Clear existing year sections
for year_section in container.find_all('div', class_='year-section'):
    year_section.decompose()

# Group the papers by year
papers_by_year = df.groupby('Year')

# Recreate the year sections
for year, papers in papers_by_year:
    year_section = soup.new_tag('div', **{'class': 'year-section'})
    
    # Year header
    year_header = soup.new_tag('h4')
    year_header.string = f'Year: {year}'
    year_section.append(year_header)
    
    # Add papers for this year
    for _, paper in papers.iterrows():
        paper_unit = soup.new_tag('div', **{'class': 'paper-unit'})
        
        # Header with title and authors
        paper_header = soup.new_tag('div', **{'class': 'paper-header'})
        header_content = soup.new_tag('div')
        header_content.append(soup.new_tag('strong'))
        header_content.contents[-1].string = 'Title: '
        header_content.append(paper['Title'])
        header_content.append(soup.new_tag('br'))
        header_content.append(soup.new_tag('strong'))
        header_content.contents[-1].string = 'Authors: '
        header_content.append(paper['Authors'])
        paper_header.append(header_content)
        
        # Expand button
        expand_button = soup.new_tag('button', **{'class': 'expand-btn', 'onclick': 'toggleDetails(this)'})
        expand_button.string = 'Show Details'
        paper_header.append(expand_button)
        paper_unit.append(paper_header)
        
        # Details section
        details = soup.new_tag('div', **{'class': 'details', 'style': 'display: none;'})
        details.append(soup.new_tag('strong'))
        details.contents[-1].string = 'Publication: '
        details.append(paper['Publication'])
        details.append(soup.new_tag('br'))
        
        details.append(soup.new_tag('strong'))
        details.contents[-1].string = 'Link: '
        link = soup.new_tag('a', href=paper['Link'])
        link.string = '[View Paper]'
        details.append(link)
        details.append(soup.new_tag('br'))
        
        details.append(soup.new_tag('strong'))
        details.contents[-1].string = 'Summary: '
        summary = soup.new_tag('p')
        summary.string = paper['Summary']
        details.append(summary)
        
        details.append(soup.new_tag('strong'))
        details.contents[-1].string = 'Main Techniques: '
        techniques = soup.new_tag('p')
        techniques.string = paper['Techniques']
        details.append(techniques)
        
        paper_unit.append(details)
        year_section.append(paper_unit)
    
    # Append the year section at the end of the container
    container.append(year_section)

# Save the updated HTML directly to the same file
with open(html_file, 'w', encoding='utf-8') as file:
    file.write(str(soup))

print(f"Updated HTML saved to {html_file}")
