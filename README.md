# Certificate Data Extractor

This project automates the process of retrieving certificate information from the [UZBMB certificate verification website](https://sertifikat.uzbmb.uz/site/cert?type=1), extracting passport and name details from the resulting PDF, and exporting the data into an Excel file.

## Features

* Automates browser interactions using **Selenium**.
* Bypasses CAPTCHA field automatically.
* Downloads certificate PDFs and parses content with **pypdf**.
* Extracts **passport number** and **full name** (ism, familiya, sharif).
* Supports multi-threaded execution for faster processing.
* Saves all results into an **Excel workbook** using **openpyxl**.

## Requirements

* Python 3.9+
* Google Chrome installed
* ChromeDriver compatible with your Chrome version

### Python dependencies:

```bash
pip install selenium requests pypdf openpyxl
```

## Usage

1. Clone the repository and navigate to the project folder.
2. Update paths for Chrome and ChromeDriver in `create_driver()`:

   ```python
   options.binary_location = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
   service = Service(r"C:\\Users\\Asus\\.cache\\selenium\\chromedriver\\win64\\128.0.6613.137\\chromedriver.exe")
   ```
3. Update the `CERT_NUMBERS` list in `pdf4.py` with the certificate numbers you want to process.
4. Run the script:

   ```bash
   python pdf4.py
   ```
5. The results will be saved into `certificates_data.xlsx`.

## Output

* The Excel file will contain three columns:

  * `cert_number`
  * `passport`
  * `full_name`

Example:

| cert\_number   | passport  | full\_name                |
| -------------- | --------- | ------------------------- |
| 24BBA1237370SR | AB1234567 | Aliyev Azizbek Anvarovich |
| 24BBA1120004ED | AC9876543 | Karimova Laylo Rahimovna  |

## Notes

* Default thread pool size is 5 (tunable in `main()`).
* If the site layout changes, element locators (`By.ID`, `By.XPATH`) may need updates.
* Avoid sending too many requests at once to prevent server blocking.
