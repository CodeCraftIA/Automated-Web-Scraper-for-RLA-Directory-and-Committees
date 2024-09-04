# Automated-Web-Scraper-for-RLA-Directory-and-Committees
A Python-based web scraper designed to extract detailed company and member information from the Reverse Logistics Association (RLA) directory and committee pages. Utilizing Selenium with undetected ChromeDriver, this script automates the process of logging in, navigating the website, and exporting the data into Excel files for further analysis.

# Overview
This project is a web scraper that automates the extraction of company profiles and committee member details from the Reverse Logistics Association (RLA) website. The scraper is built using Python and leverages Selenium with undetected ChromeDriver to bypass bot detection. The extracted data is saved into Excel files for easy access and analysis.

# Features
- Automated Login: Automatically logs into the RLA website using provided credentials.
- Company Data Extraction: Scrapes detailed information about companies listed in the RLA directory, including company overview, products, certifications, locations, and more.
- Member Data Extraction: Collects data on committee members, including their roles and associated companies.
- Excel Export: Saves the scraped data into Excel files for further analysis or reporting.

# Requirements:
- undetected_chromedriver
- tqdm
- pandas
- xlsxwriter

# Set up your credentials:
- Update the email and password variables in the with_login.py script with your RLA login credentials.
