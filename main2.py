import streamlit as st
import tempfile
import os
from combineDoctoDoc import convert_docx_to_md_and_extract_media ,redefine_from_markdown ,process_text_with_prompt,markdown_to_docx

# Assuming you have the following functions from the previous discussions:
# - convert_docx_to_md_and_extract_media
# - redefine_from_markdown
# - process_text_with_prompt
# - markdown_to_docx

# These functions need to be imported here or defined in this script

def main():
    left_co, cent_co,last_co = st.columns(3)
    with cent_co:
        st.image("C:/Users/Aditya/Downloads/Doc/revent.ai_logo.png", width=200)

    # st.title("Document Converter By Revent.ai")
    st.write("This Model in Development by Revent improves our everyday business documents. Right from formatting, to flow and structure of information and several other parameters to ensure consistency in a document are all covered. Simply test it out by uploading any word document here:")

    uploaded_file = st.file_uploader("Choose a DOCX file", type=["docx"])

    if uploaded_file is not None:
        # Button to start the conversion
        if st.button('Start Conversion'):
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_input_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_input_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())

                media_path = os.path.join(temp_dir, "media")
                temp_output_path = os.path.join(temp_dir, "final_output.docx")

                with st.spinner('Converting the document...'):
                    try:
                        # Step 1: Convert DOCX to Markdown and extract media
                        md_path = convert_docx_to_md_and_extract_media(temp_input_path, media_path)

                        # Step 2: Process the Markdown content
                        with open(md_path, "r", encoding="utf-8") as md_file:
                            markdown_content = md_file.read()
                        uniform_markdown = redefine_from_markdown(markdown_content)
                        structured_content = process_text_with_prompt(uniform_markdown)

                        # Step 3: Convert the processed content back to DOCX
                        markdown_to_docx(structured_content, temp_output_path)

                        # Offer the converted file for download
                        with open(temp_output_path, "rb") as final_docx:
                            st.download_button("Download Converted Document", final_docx, "converted_output.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    except Exception as e:
                        st.error("An error occurred during conversion.")
                        st.exception(e)

if __name__ == "__main__":
    main()