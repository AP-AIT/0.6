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

# Dropdown to select the type of data
data_type = st.selectbox("Select the type of data", ["Text", "Image", "Excel", "PDF", "Word"])

# Function to extract text from different data types
def extract_text(data_type, content):
    if data_type == "Image":
        # Extract text from image using pytesseract
        image = Image.open(io.BytesIO(content))
        text = pytesseract.image_to_string(image)
        return text
    elif data_type == "PDF":
        # Extract text from PDF
        pdf_reader = PyPDF2.PdfReader(io.BytesIO(content))
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            text += pdf_reader.pages[page_num].extract_text()
        return text
    elif data_type == "Word":
        # Extract text from Word document
        doc = Document(io.BytesIO(content))
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text
    else:
        # For other data types (Text and Excel), return the content as is
        return content

# Function to fetch email content
def fetch_email_content(username, password, search_email):
    try:
        # Connect to the local IMAP server (replace 'localhost' and '993' with your server details)
        mail = imaplib.IMAP4_SSL('localhost', 993)
        
        # Login to the email account
        mail.login(username, password)

        # Select the mailbox (e.g., 'inbox')
        mail.select('inbox')

        # Search for emails from the specified email address
        status, messages = mail.search(None, f'(FROM "{search_email}")')

        # Get the latest email ID
        latest_email_id = messages[0].split()[-1]

        # Fetch the content of the latest email
        status, msg_data = mail.fetch(latest_email_id, '(RFC822)')

        # Close the connection
        mail.logout()

        # Return the email content
        return msg_data[0][1]

    except Exception as e:
        raise Exception(f"Error connecting to the IMAP server: {e}")

# Fetch email content based on user input
if st.button("Fetch Email"):
    try:
        # Fetch email content
        email_content = fetch_email_content(user, password, search_email)

        # Extract text based on the selected data type
        extracted_text = extract_text(data_type, email_content)

        # Display the extracted text
        st.text_area("Extracted Text", extracted_text)

    except Exception as e:
        st.error(f"An error occurred: {e}")
