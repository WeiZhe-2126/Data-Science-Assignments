import pdfkit
import requests
from bs4 import BeautifulSoup
import os

# Base URL of the manual
BASE_URL = "https://cs.wingarc.com/manual/mb/6.4/en/"
START_PAGE = "UUID-0255580f-1c8e-2728-b018-5831c1cff87b.html"

# To track visited links and avoid duplicates
visited_links = set()

def fetch_links(url):
    """Fetches all valid internal links from the given URL."""
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")
    
    # Find all navigation and in-page links
    links = [
        a['href'] for a in soup.find_all('a', href=True)
        if a['href'].startswith("UUID") and a['href'] not in visited_links
    ]
    return links

def save_page_as_pdf(url, output_dir):
    """Converts a webpage to PDF and saves it."""
    pdf_file = os.path.join(output_dir, url.split("/")[-1] + ".pdf")
    pdfkit.from_url(BASE_URL + url, pdf_file)
    print(f"Saved: {pdf_file}")

def crawl_and_convert(start_page, output_dir):
    """Crawls all linked pages starting from the given page and converts them to PDFs."""
    to_visit = [start_page]
    
    while to_visit:
        current_page = to_visit.pop(0)
        if current_page in visited_links:
            continue
        
        print(f"Processing: {current_page}")
        visited_links.add(current_page)
        
        # Save current page as PDF
        save_page_as_pdf(current_page, output_dir)
        
        # Fetch new links and add to the to_visit queue
        new_links = fetch_links(BASE_URL + current_page)
        to_visit.extend(new_links)

# Directory to save PDFs
output_directory = "manual_pdfs"
os.makedirs(output_directory, exist_ok=True)

# Start crawling and converting
crawl_and_convert(START_PAGE, output_directory)
