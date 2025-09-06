from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import requests
from pypdf import PdfReader
import io
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook

# Function to configure and create a WebDriver instance
def create_driver():
    options = webdriver.ChromeOptions()
    options.binary_location = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    options.add_argument('--headless')

    service = Service(r"C:\\Users\\Asus\\.cache\\selenium\\chromedriver\\win64\\128.0.6613.137\\chromedriver.exe")
    return webdriver.Chrome(service=service, options=options)

# Function to read PDF content and combine ism, familiya, and sh
def read_pdf(content) -> dict:
    reader = PdfReader(io.BytesIO(content))
    page = reader.pages[0]
    text = page.extract_text()
    text = text.split('\n')
    
    # Combine ism, familiya, and sh into a single string
    full_name = f"{text[2]} {text[3]} {text[4]}"
    
    return {
        'passport': text[1],
        'full_name': full_name
    }

# Main function to process each certificate
def process_certificate(url, cert_number):
    driver = create_driver()
    driver.get(url)

    time.sleep(5)
    input_field = driver.find_element(By.ID, 'data-cert_number')
    input_field.send_keys(cert_number)

    # Remove the captcha field
    js_code = """
    var element = document.getElementsByClassName('form-group field-data-verifycode')[0]
    element.remove()
    """
    driver.execute_script(js_code)

    btn = driver.find_element(By.ID, 'save-see-form')
    btn.click()

    download_link = driver.find_element(By.XPATH, '//a[@class="btn btn-primary"]')
    pdf_url = download_link.get_attribute('href')
    driver.quit()
    response = requests.get(pdf_url)

    info = read_pdf(response.content)
    info['cert_number'] = cert_number  # Include certificate number for reference
    return info

# Wrapper function to call from the thread pool
def thread_main(cert_number):
    return process_certificate(SITE_URL, cert_number)

# Main function to execute multiple threads and save results to Excel
def main(cert_numbers, output_excel):
    with ThreadPoolExecutor(max_workers=5) as executor:
        results = executor.map(thread_main, cert_numbers)

    # Create an Excel workbook and worksheet
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Certificate Data"
    
    # Write headers
    headers = ['cert_number', 'passport', 'full_name']
    worksheet.append(headers)
    
    # Write data to the worksheet
    for result in results:
        worksheet.append([result['cert_number'], result['passport'], result['full_name']])
    
    # Save the workbook to a file
    workbook.save(output_excel)
    
    print(f"Data has been saved to {output_excel}")

# Example usage
SITE_URL = 'https://sertifikat.uzbmb.uz/site/cert?type=1'
CERT_NUMBERS = [
'24BBA1237370SR',
'24BBA1120004ED',
]   # Add more, certificate numbers as needed
OUTPUT_EXCEL = 'certificates_data.xlsx'

# Execute and save to Excel
main(CERT_NUMBERS, OUTPUT_EXCEL)
