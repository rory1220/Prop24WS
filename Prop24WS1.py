import imaplib
import email
import re
import os
from urllib.parse import urlparse
from selenium import webdriver
from selenium.webdriver.common.by import By
import csv
import time
import tkinter as tk
from tkinter import simpledialog
from datetime import datetime, timedelta

def get_imap_credentials():
    # Prompt
    root = tk.Tk()
    root.withdraw()
    imap_server = simpledialog.askstring(title="IMAP Login Details",
                                         prompt="Enter the IMAP server name:")
    imap_port = 993
    imap_user = simpledialog.askstring(title="IMAP Login Details",
                                       prompt="Enter your email address:")
    imap_password = simpledialog.askstring(title="IMAP Login Details",
                                           prompt="Enter your password:",
                                           show='*')
    return imap_server, imap_port, imap_user, imap_password


# Get creds
imap_server, imap_port, imap_user, imap_password = get_imap_credentials()

imap = imaplib.IMAP4_SSL(imap_server, imap_port)
imap.login(imap_user, imap_password)

mailbox = 'INBOX'
imap.select(mailbox)

# Current date?
today = datetime.now().date()
start_of_week = today - timedelta(days=today.weekday())
end_of_week = start_of_week + timedelta(days=6)

# start/end dates to str
start_of_week_str = start_of_week.strftime('%d-%b-%Y')
end_of_week_str = end_of_week.strftime('%d-%b-%Y')

# Email search
search_criteria = f'FROM "no-reply@property24.com" SINCE "{start_of_week_str}" BEFORE "{end_of_week_str}"'
print(f"Searching for emails with criteria: {search_criteria}")
result, data = imap.search(None, search_criteria)

if data == [b'']:
    print("No emails matching the search criteria were found.")
else:
    # Loop
    print(f"Fetching {len(data[0].split())} emails.")
    processed_urls = []
    for email_id in data[0].split():
        result, email_data = imap.fetch(email_id, "(RFC822)")
        raw_email = email_data[0][1]
        email_message = email.message_from_bytes(raw_email)

        # Get URLs
        urls = []
        for part in email_message.walk():
            if part.get_content_type() == 'text/html':
                html_body = part.get_payload(decode=True).decode()
                urls.extend(re.findall(r'<a\s+(?:[^>]*?\s+)?href="([^"]*)"',html_body))

        property24_urls = []
        for url in urls:
            parsed_url = urlparse(url)
            if parsed_url.netloc == 'www.property24.com' and \
                    parsed_url.path == '/General/RedirectToListing' and \
                    url not in processed_urls:
                property24_urls.append(url)
                processed_urls.append(url)

        for property24_url in property24_urls:
            options = webdriver.ChromeOptions()
            options.headless = True
            driver = webdriver.Chrome(options=options)
            driver.get(property24_url)

            try:
                price = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div/div[3]/div[1]/div[1]/div[2]').text.strip()
            except:
                print("Could not find price on the webpage.")
                price = None

            try:
                address = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div/div[3]/div[1]/div[3]/a').text.strip()
            except:
                print("Could not find address on the webpage.")
                address = None

            try:
                size = driver.find_element(By.XPATH, '//*[@id="js_accordion_propertyoverview"]/div/div[5]/div[2]/div').text.strip()
            except:
                print("Could not find size on the webpage.")
                size = None

            try:
                property_type = driver.find_element(By.XPATH, '//*[@id="js_accordion_propertyoverview"]/div/div[2]/div[2]/div').text.strip()
            except:
                print("Could not find type on the webpage.")
                property_type = None

            try:
                date_of_listing = driver.find_element(By.XPATH, '//*[@id="js_accordion_propertyoverview"]/div/div[4]/div[2]/div').text.strip()
            except:
                print("Could not find date of listing on the webpage.")
                property_type = None

             #CSV
            parsed_url = urlparse(property24_url)
            domain = parsed_url.netloc
            output_filename = f'{domain}.csv'

            # Column names
            headers = ['Price', 'Address', 'URL', 'Size', 'Property type', 'Date of Listing']

            # Check if the file exists
            file_exists = os.path.isfile(output_filename)
            with open(output_filename, mode='a', newline='') as output_file:
                csv_writer = csv.writer(output_file)
                if not file_exists:
                    csv_writer.writerow(headers)

                # Write to csv
                csv_writer.writerow([elem if elem is not None else "" for elem in [price, address, property24_url, size, property_type, date_of_listing]])

            driver.quit()
        
        if not property24_urls:
            print("No matching URLs were found in the email.")

# Logout
imap.close()
imap.logout()
print("Done.")