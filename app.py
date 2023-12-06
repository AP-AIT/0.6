import streamlit as st
import imaplib
import email
from bs4 import BeautifulSoup
import re
import pandas as pd
import base64
from docx import Document
import PyPDF2
from PIL import Image
import io
import pytesseract

# Streamlit app title
st.title("Automate2Excel: Simplified Data Transfer")

# Create input fields for the user and password
user = st.text_input("Enter your email address")
password = st.text_input("Enter your email password", type="password")

# Create input field for the email address to search for
search_email = st.text_input("Enter the email address to search for")

# Function to extract information from HTML content
def extract_info_from_html(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    
    info = {
        "Name": None,
        "Email": None,
        "Workshop Detail": None,
        "Date": None,
        "Mobile No.": None
    }

    name_element = soup.find(string=re.compile(r'Name', re.IGNORECASE))
    if name_element:
        info["Name"] = name_element.find_next('td').get_text().strip()

    email_element = soup.find(string=re.compile(r'Email', re.IGNORECASE))
    if email_element:
        info["Email"] = email_element.find_next('td').get_text().strip()

    workshop_element = soup.find(string=re.compile(r'Workshop Detail', re.IGNORECASE))
    if workshop_element:
        info["Workshop Detail"] = workshop_element.find_next('td').get_text().strip()

    date_element = soup.find(string=re.compile(r'Date', re.IGNORECASE))
    if date_element:
        info["Date"] = date_element.find_next('td').get_text().strip()

    mobile_element = soup.find(string=re.compile(r'Mobile No\.', re.IGNORECASE))
    if mobile_element:
        info["Mobile No."] = mobile_element.find_next('td').get_text().strip()

    return info

# Function to recognize file type and perform specific actions
def process_attachment(attachment):
    file_type = attachment.get_content_type()
    file_data = attachment.get_payload(decode=True)

    if file_type == 'application/msword':
        # If it's a Word document, summarize the content
        doc = Document(io.BytesIO(file_data))
        content_summary = ""
        for paragraph in doc.paragraphs:
            content_summary += paragraph.text + "\n"
        return content_summary

    elif file_type == 'application/pdf':
        # If it's a PDF document, extract text
        pdf_reader = PyPDF2.PdfFileReader(io.BytesIO(file_data))
        text = ""
        for page_num in range(pdf_reader.numPages):
            text += pdf_reader.getPage(page_num).extractText()
        return text

    elif file_type.startswith('image/'):
        # If it's an image, extract text using OCR
        image = Image.open(io.BytesIO(file_data))
        text = pytesseract.image_to_string(image)
        return text

    else:
        return None  # For unsupported file types

if st.button("Fetch and Generate Excel"):
    try:
        # URL for IMAP connection
        imap_url = 'imap.gmail.com'

        # Connection with GMAIL using SSL
        my_mail = imaplib.IMAP4_SSL(imap_url)

        # Log in using user and password
        my_mail.login(user, password)

        # Select the Inbox to fetch messages
        my_mail.select('inbox')

        # Define the key and value for email search
        key = 'FROM'
        value = search_email  # Use the user-inputted email address to search
        _, data = my_mail.search(None, key, value)

        mail_id_list = data[0].split()

        info_list = []

        # Iterate through messages and extract information from HTML content
        for num in mail_id_list:
            typ, data = my_mail.fetch(num, '(RFC822)')
            msg = email.message_from_bytes(data[0][1])

            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    html_content = part.get_payload(decode=True).decode('utf-8')
                    info = extract_info_from_html(html_content)

                    # Extract and add the received date
                    date = msg["Date"]
                    info["Received Date"] = date

                    info_list.append(info)

                elif part.get('Content-Disposition') is not None:
                    # Process attachments
                    attachment = part
                    attachment_data = process_attachment(attachment)

                    if attachment_data:
                        st.write(f"Attachment content:\n{attachment_data}")

        # Create a DataFrame from the info_list
        df = pd.DataFrame(info_list)

        # Generate the Excel file
        st.write("Data extracted from emails:")
        st.write(df)

        if st.button("Download Excel File"):
            excel_file = df.to_excel('EXPO_leads.xlsx', index=False, engine='openpyxl')
            if excel_file:
                with open('EXPO_leads.xlsx', 'rb') as file:
                    st.download_button(
                        label="Click to download Excel file",
                        data=file,
                        key='download-excel'
                    )

        st.success("Excel file has been generated and is ready for download.")

    except Exception as e:
        st.error(f"Error: {e}")
