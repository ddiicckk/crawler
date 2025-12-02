
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches
import pandas as pd
from io import BytesIO
from PIL import Image

def urls_from_excel_to_word_with_images(excel_file, sheet_name, url_column, output_file):
    # Read URLs from Excel
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    urls = df[url_column].dropna().tolist()  # Remove empty cells

    # Create a Word document
    doc = Document()
    doc.add_heading('Web Content Compilation', level=0)

    for idx, url in enumerate(urls, start=1):
        try:
            # Fetch the webpage content
            response = requests.get(url)
            response.raise_for_status()

            # Parse HTML
            soup = BeautifulSoup(response.text, 'html.parser')

            # Extract text elements
            text_elements = soup.find_all(['h1', 'h2', 'h3', 'p', 'li'])

            # Add section for this URL
            doc.add_page_break()
            doc.add_heading(f'Source {idx}: {url}', level=1)

            for element in text_elements:
                if element.name.startswith('h'):
                    doc.add_heading(element.get_text(strip=True), level=int(element.name[1]))
                else:
                    doc.add_paragraph(element.get_text(strip=True))

            # Extract and add images
            images = soup.find_all('img')
            for img in images[:5]:  # Limit to first 5 images per page for performance
                img_url = img.get('src')
                if img_url and img_url.startswith(('http', 'https')):
                    try:
                        img_response = requests.get(img_url)
                        img_response.raise_for_status()
                        image = Image.open(BytesIO(img_response.content))
                        image.thumbnail((600, 600))  # Resize for Word
                        img_stream = BytesIO()
                        image.save(img_stream, format='PNG')
                        img_stream.seek(0)
                        doc.add_picture(img_stream, width=Inches(3))
                    except Exception as img_err:
                        doc.add_paragraph(f"Image could not be added: {img_err}")

            print(f"Processed: {url}")

        except Exception as e:
            doc.add_page_break()
            doc.add_heading(f'Source {idx}: {url}', level=1)
            doc.add_paragraph(f"Failed to retrieve content. Error: {e}")
            print(f"Error processing {url}: {e}")

    # Save the document
    doc.save(output_file)
    print(f"All content saved to {output_file}")

# Example usage:
excel_file = "urls.xlsx"       # Path to your Excel file
sheet_name = "Sheet1"          # Name of the sheet
url_column = "URL"             # Column name containing URLs
output_file = "web_content_with_images.docx"

urls_from_excel_to_word_with_images(excel_file, sheet_name, url_column, output_file)
