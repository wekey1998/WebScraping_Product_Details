#!/usr/bin/env python
# coding: utf-8

# In[1]:


##**Description**: Extracts product details (Price, Title, SKU) from a list of URLs in an Excel file.##


import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import time

# User Inputs: Provide your own file paths
excel_path = "input.xlsx"  # Path to input Excel file containing URLs
output_path = "output.xlsx"  # Path to save the extracted data
chrome_driver_path = "chromedriver.exe"  # Path to ChromeDriver

# Load the Excel file
urls_df = pd.read_excel(excel_path)

# Ensure required columns
required_columns = ["URLs", "Price", "Title", "SKU", "Loaded URL"]
for col in required_columns:
    if col not in urls_df.columns:
        urls_df[col] = ""

# Initialize WebDriver
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service)
driver.maximize_window()

# XPaths for elements
price_xpath = "//span[@class='font-size-3 font-weight-7 js-price']"
title_xpath = "//div[@class='font-size-3 letter-spacing-1 font-weight-5 line-height-copy-tall margin-bottom-3']"
sku_xpath = "//h1[@class='font-size-4 font-weight-9 letter-spacing-1 margin-bottom-3']"

# Process URLs
for index, row in urls_df.iterrows():
    url = row["URLs"]
    if pd.isna(url):
        urls_df.at[index, ["Price", "Title", "SKU", "Loaded URL"]] = "Invalid URL"
        continue

    try:
        # Navigate to the URL
        driver.get(url)
        time.sleep(3)  # Allow page to load

        # Store the loaded URL
        urls_df.at[index, "Loaded URL"] = driver.current_url

        # Extract price
        try:
            price_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, price_xpath))
            )
            urls_df.at[index, "Price"] = price_element.text.strip()
        except:
            urls_df.at[index, "Price"] = "Price Not Found"

        # Extract title
        try:
            title_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, title_xpath))
            )
            urls_df.at[index, "Title"] = title_element.text.strip()
        except:
            urls_df.at[index, "Title"] = "Title Not Found"

        # Extract SKU
        try:
            sku_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, sku_xpath))
            )
            urls_df.at[index, "SKU"] = sku_element.text.strip()
        except:
            urls_df.at[index, "SKU"] = "SKU Not Found"

    except Exception as e:
        print(f"Error processing URL at index {index}: {e}")
        urls_df.at[index, ["Price", "Title", "SKU", "Loaded URL"]] = "Error"

# Save the updated Excel file
urls_df.to_excel(output_path, index=False)

# Quit the driver
driver.quit()
print(f"Data has been extracted and saved to {output_path}.")


