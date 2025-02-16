# WebScraping_Product_Details
This automation extracts product details (Price, Title, and SKU) from a list of URLs provided in an Excel file. The script uses Selenium to navigate each URL, scrape the required details, and save the results into a new Excel file.
# Web Scraping Product Details using Selenium

## Overview
This project automates the extraction of product details (Price, Title, and SKU) from a list of URLs provided in an Excel file. The script uses Selenium to navigate each URL, scrape the required details, and save the results into a new Excel file.

## Features
- Reads product URLs from an input Excel file.
- Uses Selenium to extract product details.
- Saves the extracted data into an output Excel file.
- Handles missing elements and errors gracefully.

## Technologies Used
- Python
- Selenium
- Pandas
- OpenPyXL

## Installation
### Prerequisites
Ensure you have the following installed on your system:
- [Python 3.x](https://www.python.org/downloads/)
- [Google Chrome](https://www.google.com/chrome/)
- [ChromeDriver](https://sites.google.com/chromium.org/driver/)

### Install Dependencies
Run the following command to install the required Python libraries:
```sh
pip install pandas selenium openpyxl
```

## Usage
1. **Prepare Input File:**
   - Create an Excel file (`input.xlsx`) with a column named `URLs` containing the product URLs.

2. **Set File Paths:**
   - Update the script with your file paths:
     ```python
     excel_path = "input.xlsx"  # Path to input file
     output_path = "output.xlsx"  # Path to save extracted data
     chrome_driver_path = "chromedriver.exe"  # Path to ChromeDriver
     ```

3. **Run the Script:**
   ```sh
   python script.py
   ```

4. **Check Output:**
   - The extracted product details will be saved in `output.xlsx`.

## Script Details
- **Navigates to each product URL.**
- **Extracts the following details:**
  - **Price** (`//span[@class='font-size-3 font-weight-7 js-price']`)
  - **Title** (`//div[@class='font-size-3 letter-spacing-1 font-weight-5 line-height-copy-tall margin-bottom-3']`)
  - **SKU** (`//h1[@class='font-size-4 font-weight-9 letter-spacing-1 margin-bottom-3']`)
- **Handles errors** if elements are missing.

## Troubleshooting
- **Chromedriver version mismatch:** Ensure your ChromeDriver version matches your Chrome browser version.
- **Web elements not found:** Verify the XPath selectors are correct for the website you're scraping.
- **Excel file issues:** Ensure `input.xlsx` has a `URLs` column.

## Author
**Vigneshwaran**

