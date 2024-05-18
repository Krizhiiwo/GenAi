import streamlit as st
import tempfile
import docx
import os
import openai
from dotenv import load_dotenv
from io import BytesIO

def load_variable():
    load_dotenv('.venv/key')
    return os.environ.get('OPENAI_API_KEY')

def text_extract(file):
    doc_text = ""
    if file is not None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_file_path = f"{temp_dir}/uploaded_file.docx"

            with open(temp_file_path, "wb") as temp_file:
                temp_file.write(file.read())

            doc = docx.Document(temp_file_path)
            for para in doc.paragraphs:
                doc_text += para.text + "\n"
    return doc_text

def main():
    openai_api_key = load_variable()
    st.title("Doc Text Extractor")

    uploaded_file = st.file_uploader("Upload a DOCX file", type=["docx"])

    if uploaded_file is not None:
        doc_contents = text_extract(uploaded_file)
        st.header("DOCX Contents:")
        st.text(doc_contents)

        if openai_api_key:
            openai.api_key = openai_api_key  # Set the OpenAI API key
            if st.button("Process with OpenAI"):
                try:
                    response = openai.Completion.create(
                        model="gpt-3.5-turbo",
                        prompt=f"Extracted Text: {doc_contents}\n\nProvide a summary:",
                        max_tokens=150
                    )
                    st.header("OpenAI Response:")
                    st.text(response.choices[0].text.strip())
                except Exception as e:
                    st.error(f"Error using OpenAI API: {e}")
        else:
            st.error("OpenAI API key is not set")

if __name__ == "__main__":
    main()
