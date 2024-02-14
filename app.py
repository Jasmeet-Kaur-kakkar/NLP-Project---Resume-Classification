# -*- coding: utf-8 -*-
"""
Created on Mon Jan 29 20:43:59 2024

"""

import streamlit as st
import os
import re
from pickle import load
import docx
from PyPDF2 import PdfReader
import docx2txt
import pandas as pd
import pythoncom
from win32com import client

# Load the pre-trained model, TF-IDF Vectorizer, and LabelEncoder
model = load(open('model.pkl', 'rb'))
vectorizer = load(open('vectorizer.pkl', 'rb'))
encoder = load(open('encoder.pkl', 'rb'))

# Function to preprocess a resume
def preprocess_resume(file):
    file_extension = os.path.splitext(file.name)[1]

    if file_extension.lower() == ".pdf":
        pdf_reader = PdfReader(file)
        content = ""
        for page_num in range(len(pdf_reader.pages)):
            content += pdf_reader.pages[page_num].extract_text()
    elif file_extension.lower() == ".docx":
        try:
            doc = docx.Document(file)
            content = ""
            for para in doc.paragraphs:
                content += para.text + "\n"
        except Exception as e:
            st.error(f"Error reading .docx file: {e}")
            return None
    elif file_extension.lower() == ".doc":
        try:
            pythoncom.CoInitialize()
            word = client.Dispatch("Word.Application")
            doc = word.Documents.Open(file.name)
            content = doc.Content.Text
            doc.Close()
            word.Quit()
        except Exception as e:
            st.error(f"Error reading .doc file: {e}")
            return None
    else:
        st.error("Unsupported file format. Please upload a .doc, .docx, or .pdf file.")
        return None
    
    # Remove non-ASCII characters (excluding English alphabets and common symbols)
    content = re.sub(r'[^\x00-\x7F]+', ' ', content)
    # Remove control characters (ASCII values 0-31, excluding tab, newline, and carriage return)
    content = re.sub(r'[\x01-\x08\x0B\x0C\x0E-\x1F]+', ' ', content)
    # Remove control characters 
    content = ''.join(char for char in content if ord(char) > 31 or ord(char) == 9 or ord(char) == 10 or ord(char) == 13)
    # Remove hashtags
    content = re.sub(r'#\S+', '', content, flags=re.MULTILINE)
    # Remove special characters and numbers
    content = re.sub(r'[^a-zA-Z\s]', '', content)
    # Remove mentions (@username)
    content = re.sub(r'@\S+', '', content, flags=re.MULTILINE)
    # Remove URLs
    content = re.sub(r'http\S+|www\S+|https\S+', '', content, flags=re.MULTILINE)
    # Remove RT and CC
    content = re.sub(r'\bRT\b|\bCC\b', '', content, flags=re.MULTILINE)

    return content

# Function to get list of files in a folder
def list_files(folder_path):
    files = []
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            files.append(os.path.join(dirpath, filename))
    return files

# Streamlit App
def main():
    st.title("Resume Category Predictor")

    # Allow user to input folder path
    folder_path = st.text_input("Enter the path to the folder containing resumes:", key="folder_path")

    if folder_path and os.path.exists(folder_path):
        # List all files in the selected folder
        files = list_files(folder_path)

        # Display the list of files
        selected_files = st.multiselect("Select resumes for prediction", files)

        if selected_files:
            for file_path in selected_files:
                st.subheader(f"Resume Preview - {os.path.basename(file_path)}:")
                
                # Read the content of the file
                with open(file_path, 'rb') as file:
                    content = preprocess_resume(file)
                    if content:
                        st.text_area(f"Resume Content - {os.path.basename(file_path)}", content, height=250)

                        # Apply TF-IDF Vectorization
                        tfidf_matrix = vectorizer.transform([content])
                        tfidf_df = pd.DataFrame(tfidf_matrix.toarray(), columns=vectorizer.get_feature_names_out())

                        # Make Prediction
                        prediction = model.predict(tfidf_df)
                        predicted_category = encoder.inverse_transform(prediction)[0]

                        st.subheader(f"Resume Category - {os.path.basename(file_path)}:")
                        st.success(f"Classified Category: {predicted_category}")


if __name__ == "__main__":
    main()