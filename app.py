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

# Fetch email content based on user input
if st.button("Fetch Email"):
    # Add logic to fetch email content here based on user input
    # Example: use IMAP to connect to the email server, fetch the email, and extract content

    # For demonstration purposes, I'll assume you have the email content and type in variables content and content_type
    content = b"Your email content as bytes"  # Replace with actual email content
    content_type = "application/pdf"  # Replace with actual content type

    # Extract text based on the selected data type
    extracted_text = extract_text(data_type, content)

    # Display the extracted text
    st.text_area("Extracted Text", extracted_text)
