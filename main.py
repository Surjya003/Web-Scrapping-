import requests
from bs4 import BeautifulSoup
import pandas as pd
from collections import Counter

# List of companies with their URLs and categories
companies = [
    {"name": "Nestle", "url": "https://www.nestle.com", "category": "F&B", "Manufacturer": "No", "Brand": "No", "Distributor": "No"},
    {"name": "Pfizer", "url": "https://www.pfizer.com", "category": "Distributor", "Manufacturer": "No", "Brand": "No", "Distributor": "Yes"},
    {"name": "Johnson & Johnson", "url": "https://www.jnj.com", "category": "Manufacturer", "Manufacturer": "Yes", "Brand": "No", "Distributor": "No"},
    {"name": "Unilever", "url": "https://www.unilever.com", "category": "F&B", "Manufacturer": "No", "Brand": "No", "Distributor": "No"},
    {"name": "General Mills", "url": "https://www.generalmills.com", "category": "Manufacturer", "Manufacturer": "Yes", "Brand": "No", "Distributor": "No"},
    {"name": "Kellogg’s", "url": "https://www.kelloggs.com", "category": "Manufacturer", "Manufacturer": "Yes", "Brand": "No", "Distributor": "No"}
]

# Keywords to search for
keywords = {
    "Probiotics": ["probiotic", "probiotics"],
    "Fortification": ["fortified", "fortification"],
    "Gut Health": ["gut health", "digestive health"],
    "Women’s Health": ["women's health", "female health", "PCOD", "UTI"],
    "Cognitive Health": ["cognitive health", "mental wellness", "anxiety"]
}

# Add verticals and product details
verticals = {
    "F&B": "Food & Beverages: Drinks, Milk, Cereal, Bakery products.",
    "Bulk": "Pharma/Nutra companies that manufacture probiotic products using UBL's strains.",
    "Formulations": "End product sold to brands that are into relevant health segments."
}

examples = {
    "F&B": "You can pitch Bacillus Coagulans Unique IS2 to Nestle Milk, Pulpy Orange.",
    "Bulk": "Manufacturers like Dr.Reddys can add our strains like Bacillus subtilis in their Gut Health products.",
    "Formulations": "Brands like Ferment ISKO which is into GutHealth can market UBL's Bacipro which is in the Gut Health space."
}

ubl_health_segments = {
    "Gut Health": "Helps with constipation, diarrhea, digestive problems.",
    "Women’s Health": "Helps with UTI, PCOD.",
    "Cognitive Health": "Reduces anxiety and enhances mental wellness."
}

# Scrape websites for keywords
def scrape_website(url, keywords):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    try:
        response = requests.get(url, headers=headers, timeout=30)  # Increased timeout
        response.raise_for_status()  # Check if request was successful
        soup = BeautifulSoup(response.text, "html.parser")
        text = soup.get_text().lower()  # Get all text from the website and convert to lowercase
        
        found_keywords = {key: any(keyword.lower() in text for keyword in kw_list) for key, kw_list in keywords.items()}
        return found_keywords
    except requests.RequestException as e:
        print(f"Error scraping {url}: {e}")
        return {key: False for key in keywords.keys()}

# Prepare data for both Excel sheets
csv_data = []
scraped_data = []

for company in companies:
    company_name = company["name"]
    url = company["url"]
    category = company["category"]
    manufacturer = company["Manufacturer"]
    brand = company["Brand"]
    distributor = company["Distributor"]
    
    # Scrape website for keywords
    result = scrape_website(url, keywords)
    
    # Determine presence of keywords
    keyword_results = {key: 'Yes' if result[key] else 'No' for key in keywords.keys()}
    
    # Determine category explanation
    category_explanation = verticals.get(category, "No explanation available.")
    
    # Prepare CSV data
    csv_data.append({
        "Company Name": company_name,
        "URL": url,
        "Category": category,
        "Manufacturer": manufacturer,
        "Brand": brand,
        "Distributor": distributor,
        "Category Explanation": category_explanation,
        **keyword_results
    })
    
    # Prepare Scraped Data
    scraped_data.append({
        "Company Name": company_name,
        "Website": url,
        "Probiotics": result["Probiotics"],
        "Fortification": result["Fortification"],
        "Gut Health": result["Gut Health"],
        "Women’s Health": result["Women’s Health"],
        "Cognitive Health": result["Cognitive Health"]
    })

# Convert data to DataFrames
df_keywords = pd.DataFrame(csv_data)
df_scraped = pd.DataFrame(scraped_data)

# Save both DataFrames to a single Excel file with multiple sheets
output_file = "company_data_with_details.xlsx"
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    df_keywords.to_excel(writer, sheet_name='Company Keywords', index=False)
    df_scraped.to_excel(writer, sheet_name='Scraped Data', index=False)
    
    # Access the XlsxWriter workbook and worksheets
    workbook  = writer.book
    worksheet_keywords = writer.sheets['Company Keywords']
    
    # Add a new column for the Prospect status
    worksheet_keywords.write('I1', 'Prospect')
    
    # Apply a formula to categorize based on Manufacturer
    worksheet_keywords.write_formula('I2', '=IF(D2="Yes","Prospect","Not Relevant")', workbook.add_format({'bg_color': '#FFFF00'}))
    
    # Apply conditional formatting
    worksheet_keywords.conditional_format('H2:H{}'.format(len(df_keywords) + 1),
                                          {'type': 'text',
                                           'criteria': 'containing',
                                           'value': 'Gut Health',
                                           'format': workbook.add_format({'bg_color': '#FFCCCC'})})

    worksheet_keywords.conditional_format('H2:H{}'.format(len(df_keywords) + 1),
                                          {'type': 'text',
                                           'criteria': 'containing',
                                           'value': 'Women’s Health',
                                           'format': workbook.add_format({'bg_color': '#CCFFCC'})})

    worksheet_keywords.conditional_format('H2:H{}'.format(len(df_keywords) + 1),
                                          {'type': 'text',
                                           'criteria': 'containing',
                                           'value': 'Cognitive Health',
                                           'format': workbook.add_format({'bg_color': '#CCCCFF'})})

print(f"Data saved to {output_file}")
