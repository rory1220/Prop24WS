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

# Define the GUI
root = tk.Tk()
root.title("Property24 Scraper")
root.geometry("400x400")

# Define the status label
status_label = tk.Label(root, text="")
status_label.pack()

# Define the start button
def start_scraping():

    # Prompt user for IMAP server, login credentials
    imap_server = input("Enter IMAP server address: ")
    imap_port = input("Enter IMAP Port: ")
    imap_user = input("Enter email address: ")
    imap_password = input("Enter email password: ")

    # Define the email search criteria
    search_criteria = 'FROM "no-reply@property24.com"'
    print(f"Searching for emails with criteria: {search_criteria}")
    result, data = imap.search(None, search_criteria)
    
    try:
        # Connect to the IMAP server and login
        imap = imaplib.IMAP4_SSL(imap_server, imap_port)
        imap.login(imap_user, imap_password)
        
        # Select the inbox
        imap.select('INBOX')
        
        # Search for emails matching search criteria
        result, data = imap.search(None, search_criteria)
        
        # Display message box with number of emails found
        messagebox.showinfo(title="Search Complete", message=f"{len(data[0].split())} emails found!")
        
        # Logout of the IMAP server
        imap.logout()
    except imaplib.IMAP4.error as e:
        messagebox.showerror(title="Error", message=f"Error: {e}")

    # Search for emails matching the specified criteria
    search_criteria = f'FROM "{search_entry.get()}"'
    status_label.config(text=f"Searching for emails with criteria: {search_criteria}")
    result, data = imap.search(None, search_criteria)

    if data == [b'']:
        status_label.config(text="No emails matching the search criteria were found.")
    else:
        # Loop through the list of email IDs returned by the search and fetch each email
        status_label.config(text=f"Fetching {len(data[0].split())} emails.")
        processed_urls = []
        for email_id in data[0].split():
            result, email_data = imap.fetch(email_id, "(RFC822)")
            raw_email = email_data[0][1]
            email_message = email.message_from_bytes(raw_email)

            # Extract the URLs from the email message
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
                # Extract the relevant data from the Property24 website
                options = webdriver.ChromeOptions()
                options.headless = True
                driver = webdriver.Chrome(options=options)
                driver.get(property24_url)

                try:
                    price = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div/div[3]/div[1]/div[1]/div[2]').text.strip()
                except:
                    status_label.config(text="Could not find price on the webpage.")
                    price = "Could not find price on the webpage."

                try:
                    address = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div/div[3]/div[1]/div[3]/a').text.strip()
                except:
                    status_label.config(text="Could not find address on the webpage.")
                    address = "Could not find address on the webpage."
                
                try:
                    suburb = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div/div[3]/div[1]/div[3]/span[1]').text.strip()
                except:
                    status_label.config(text="Could not find suburb on the webpage.")
                    suburb = "Could not find suburb on the webpage."

                try:
                    bedrooms = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div/div[4]/div[1]/div[2]').text.strip()
                except:
                    status_label.config(text="Could not find number of bedrooms on the webpage.")
                    bedrooms = "Could not find number of bedrooms on the webpage."

                try:
                    bathrooms = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div/div[4]/div[2]/div[2]').text.strip()
                except:
                    status_label.config(text="Could not find number of bathrooms on the webpage.")
                    bathrooms = "Could not find number of bathrooms on the webpage."

                try:
                    garages = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div/div[4]/div[3]/div[2]').text.strip()
                except:
                    status_label.config(text="Could not find number of garages on the webpage.")
                    garages = "Could not find number of garages on the webpage."

                try:
                    erf_size = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div/div[4]/div[4]/div[2]').text.strip()
                except:
                    status_label.config(text="Could not find erf size on the webpage.")
                    erf_size = "Could not find erf size on the webpage."

                try:
                    floor_size = driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[2]/div/div/div[4]/div[5]/div[2]').text.strip()
                except:
                    status_label.config(text="Could not find floor size on the webpage.")
                    floor_size = "Could not find floor size on the webpage."

            # Close the webdriver
            driver.quit()

            # Write the data to a CSV file
            with open('property24_data.csv', mode='a', newline='') as csv_file:
                fieldnames = ['URL', 'Price', 'Address', 'Suburb', 'Bedrooms', 'Bathrooms', 'Garages', 'Erf Size', 'Floor Size']
                writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                if csv_file.tell() == 0:
                    writer.writeheader()
                writer.writerow({'URL': property24_url, 'Price': price, 'Address': address, 'Suburb': suburb, 'Bedrooms': bedrooms, 'Bathrooms': bathrooms, 'Garages': garages, 'Erf Size': erf_size, 'Floor Size': floor_size})

    status_label.config(text=f"Finished fetching {len(data[0].split())} emails.")
    imap.close()
    imap.logout()

start_button = tk.Button(root, text="Start", command=start_scraping)
start_button.pack()

root.mainloop()
