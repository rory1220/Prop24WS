# Email Scraper and Property Data Extractor

## Overview

This Python script connects to an IMAP email server, searches for emails from a specific sender within the current week, extracts Property24 URLs from the emails, and scrapes data from the corresponding web pages. The extracted data includes property details such as price, address, size, property type, and date of listing.

## Prerequisites

Ensure you have the following installed:

- Python 3.x
- Chrome browser
- [ChromeDriver](https://sites.google.com/chromium.org/driver/) (or use `chromedriver-autoinstaller`)

## Installation

1. Clone. Obviously.

2. Install the required Python packages:

pip install -r requirements.txt

3. Ensure you have the Chrome browser installed.

(Optional) If not using chromedriver-autoinstaller, download the appropriate version of ChromeDriver and place it in the project directory.

## Usage

1. Run the script:

python property_scraper.py

2. Follow the prompts to enter your IMAP server details and login credentials.

3. The script will search for emails, extract Property24 URLs, scrape data, and save it to CSV files.

## Notes
The script uses a headless Chrome browser for web scraping.
Make sure your email account allows IMAP access.
Ensure your Chrome browser is up-to-date.
