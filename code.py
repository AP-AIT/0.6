import streamlit as st
import imaplib
import email
from bs4 import BeautifulSoup
import re
import pandas as pd
from PIL import Image
import io
import docx2txt
import openpyxl
import PyPDF2
import pytesseract

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

    # ... (existing code for extracting information from HTML)

    return info

# Function to recognize file type and perform specific actions
def process_attachment(attachment):
    file_type = attachment.get_content_type()
    file_data = attachment.get_payload(decode=True)

    if file_type == 'application/msword':
        # ... (existing code for processing Word documents)

    elif file_type == 'application/pdf':
        # ... (existing code for processing PDF documents)

    elif file_type.startswith('image/'):
        # ... (existing code for processing images)

    else:
        return None  # For unsupported file types

# ... (other functions and imports)

# Streamlit app title
st.title("Automate2Excel: Simplified Data Transfer")

# Create input fields for the user and password
user = st.text_input("Enter your email address")
password = st.text_input("Enter your email password", type="password")

# Create input field for the email address to search for
search_email = st.text_input("Enter the email address to search for")

# ... (other input fields and functions)

# Initialize mail_id_list outside the try block
mail_id_list = []

# Define a variable to store the selected document type
document_type = st.selectbox("Select Document Type", ["HTML", "Image", "Word", "Excel", "PDF"])

if st.button("Fetch and Generate Extracted Document"):
    try:
        # URL for IMAP connection
        imap_url = 'imap.gmail.com'

        # Connection with Gmail using SSL
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

        # Iterate through messages and extract information based on the selected document type
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

                elif document_type == "Image" and part.get_content_type().startswith("image"):
                    # Call the function to extract text from image
                    image_text = extract_text_from_image(part.get_payload(decode=True))
                    st.text(f"Extracted Text from Image: {image_text}")

                elif document_type == "Word" and part.get_content_type().startswith("application/msword"):
                    # Call the function to extract text from Word document
                    word_text = extract_text_from_word(part.get_payload(decode=True))
                    st.text(f"Extracted Text from Word: {word_text}")

                elif document_type == "Excel" and part.get_content_type().startswith("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
                    # Call the function to extract text from Excel file
                    excel_text = extract_text_from_excel(part.get_payload(decode=True))
                    st.text(f"Extracted Text from Excel: {excel_text}")

                elif document_type == "PDF" and part.get_content_type().startswith("application/pdf"):
                    # Call the function to extract text from PDF
                    pdf_text = extract_text_from_pdf(part.get_payload(decode=True))
                    st.text(f"Extracted Text from PDF: {pdf_text}")

        # Create a DataFrame from the info_list
        df = pd.DataFrame(info_list)

        # Generate the Excel file
        st.write("Data extracted from emails:")
        st.write(df)

        if st.button("Download Extracted Document"):
            # Customize the download functionality based on your requirements
            # You might want to save the extracted content to a file and provide a download link
            # Example: Save to a text file and provide a download link
            with open('extracted_content.txt', 'w') as file:
                for info in info_list:
                    file.write(str(info) + '\n')

            st.download_button(
                label="Click to download Extracted Document",
                data='extracted_content.txt',
                key='download-extracted-document'
            )

        st.success("Extraction completed.")

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.warning("Click the 'Fetch and Generate Extracted Document' button to retrieve and process emails.")
