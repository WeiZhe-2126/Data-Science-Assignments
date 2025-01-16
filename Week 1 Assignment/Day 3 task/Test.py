import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import asyncio
from pyppeteer import connect
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

async def print_all_pages_with_pyppeteer():
    chrome_options = Options()
    chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-gpu")
    driver = webdriver.Chrome(options=chrome_options)

    # Navigate using Selenium
    driver.get("https://cs.wingarc.com/manual/mb/6.4/en/UUID-0255580f-1c8e-2728-b018-5831c1cff87b.html")
    
    # Connect Pyppeteer to the existing browser
    browser = await connect(browserURL='http://127.0.0.1:9222')

    while True:
        pages = await browser.pages()
        page = pages[0]  # Assume the first page is the one to manipulate

        # Wait for 5 seconds to let the page load
        time.sleep(1)

        # Hide hyperlink URLs before printing
        await page.addStyleTag({'content': 'a::after { content: none !important; }'})

        # Get the page title and sanitize it for file naming
        page_title = await page.title()
        safe_title = "".join(c if c.isalnum() or c in (' ', '-', '_') else "_" for c in page_title).strip()
        file_name = f"/Users/weizhe/Downloads/Automation Assignment/Test2/{safe_title}.pdf"

        # Save the PDF
        await page.pdf({
            'path': file_name,
            'format': 'A4',
            'printBackground': False,
            'scale': 1,
            'fullPage': True,
            'margin': {'top': '10mm', 'bottom': '10mm', 'left': '10mm', 'right': '10mm'},
        })

        # Attempt to navigate to the next page
        try:
            next_button = driver.find_element(By.LINK_TEXT, "Next")
            next_button.click()
            time.sleep(0.5)  # Short delay to allow the next page to load
        except NoSuchElementException:
            print("No more pages. Printing completed.")
            break

    await browser.disconnect()
    driver.quit()

# Run the asynchronous task
asyncio.get_event_loop().run_until_complete(print_all_pages_with_pyppeteer())
