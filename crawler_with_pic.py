
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches
import pandas as pd
from io import BytesIO
from PIL import Image
import re
import os
from urllib.parse import urljoin

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

def urls_from_excel_main_content_with_images(excel_file, sheet_name, url_column, output_folder):
    # Read URLs from Excel
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    urls = df[url_column].dropna().tolist()

    os.makedirs(output_folder, exist_ok=True)

    for idx, url in enumerate(urls, start=1):
        try:
            response = requests.get(url)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')

            # Get page title for filename
            title_tag = soup.find('title')
            page_title = title_tag.get_text(strip=True) if title_tag else f"Page_{idx}"
            file_name = sanitize_filename(page_title[:100]) + ".docx"
            file_path = os.path.join(output_folder, file_name)

            # Find main content block
            main_content = soup.find('main') or soup.find('article') or soup.find('div', class_=re.compile(r'(content|article|post|body)'))
            if not main_content:
                main_content = soup  # fallback

            doc = Document()
            doc.add_heading(page_title, level=0)
            doc.add_paragraph(f"Source URL: {url}")

            # Extract unique text
            text_elements = main_content.find_all(['h1', 'h2', 'h3', 'p', 'li'])
            seen_text = set()
            for element in text_elements:
                text = element.get_text(strip=True)
                if text and text not in seen_text:
                    seen_text.add(text)
                    if element.name.startswith('h'):
                        doc.add_heading(text, level=int(element.name[1]))
                    else:
                        doc.add_paragraph(text)

            # Extract and add images (limit to 5)
            images = main_content.find_all('img')
            for img in images[:5]:
                img_url = img.get('src') or img.get('data-src')
                if img_url:
                    img_url = urljoin(url, img_url)  # Convert relative to absolute
                    try:
                        img_response = requests.get(img_url)
                        img_response.raise_for_status()
                        image = Image.open(BytesIO(img_response.content))
                        image.thumbnail((600, 600))
                        img_stream = BytesIO()
                        image.save(img_stream, format='PNG')
                        img_stream.seek(0)
                        doc.add_picture(img_stream, width=Inches(3))
                    except Exception as img_err:
                        doc.add_paragraph(f"Image could not be added: {img_err}")

            doc.save(file_path)
            print(f"Saved main content with images: {file_path}")

        except Exception as e:
            print(f"Error processing {url}: {e}")

    print(f"All files saved in folder: {output_folder}")

# Example usage:
excel_file = "urls.xlsx"
sheet_name = "Sheet1"
url_column = "URL"
output_folder = "main_content_pages"

urls_from_excel_main_content_with_images(excel_file, sheet_name, url_column, output_folder)
