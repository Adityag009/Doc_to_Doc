import openai
import dotenv
import os
from docx import Document
import subprocess


dotenv.load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")


# Function to convert DOCX to Markdown and extract media
def convert_docx_to_md_and_extract_media(docx_file_path, media_path):
    pandoc_path = os.getenv('PANDOC_PATH')
    if pandoc_path is None:
        raise ValueError("PANDOC_PATH environment variable is not set.")

    if not os.path.exists(media_path):
        os.makedirs(media_path)

    output_md_path = os.path.join(media_path, "temp.md")
    try:
        subprocess.run([pandoc_path, "-s", docx_file_path, "-t", "markdown",
                        "--extract-media", media_path, "-o", output_md_path], check=True)
    except subprocess.CalledProcessError as e:
        raise ValueError(f"Error during conversion: {e}")
    return output_md_path



# style_prompt = """
# Your task is to ensure uniformity in the provided document while preserving its original style and tone. Begin by closely analyzing the initial sentences and paragraphs to understand the nuances of the writing style, including the usage of adverbs, adjectives, sentence structure, and word choice. Maintain the format and structure of the initial text, including the number of paragraphs. Incorporate these style elements throughout the entire document without altering the format. 
# Adjust language and wording as necessary to maintain coherence, but ensure that the structure and format of the text remain consistent with the original. Pay careful attention to the voice and perspective established in the beginning and strive to emulate it consistently. Your objective is to create a seamless transition between the original content and the converted format while preserving the essence and personality conveyed in the text.
# Please provide the output in plain text format without any Markdown.

# """

def redefine_from_markdown(markdown_content):
    style_prompt = """
Your task is to ensure uniformity in the provided document while preserving its original style and tone. Begin by closely analyzing the initial sentences and paragraphs to understand the nuances of the writing style, including the usage of adverbs, adjectives, sentence structure, and word choice. Maintain the format and structure of the initial text, including the number of paragraphs. Incorporate these style elements throughout the entire document without altering the format. 
Adjust language and wording as necessary to maintain coherence, but ensure that the structure and format of the text remain consistent with the original. Pay careful attention to the voice and perspective established in the beginning and strive to emulate it consistently. Your objective is to create a seamless transition between the original content and the converted format while preserving the essence and personality conveyed in the text.
Please provide the output in plain text format without any Markdown.

"""
# 169 tokens style_prompt
    response = openai.chat.completions.create(model="gpt-4",
        messages=[
            {"role": "system", "content": style_prompt},
            {"role": "user", "content": markdown_content}
        ],
        max_tokens=4000,
        temperature=0.3)

    extracted_elements = response.choices[0].message.content

    if isinstance(extracted_elements, list):
        extracted_elements = " ".join(extracted_elements)

    return extracted_elements


# 309 tokens insturction prompt 

def process_text_with_prompt(text):
    """
    Process the extracted text with a structured prompt for formatting and style adjustments.
    """
    instructional_part = """Instructions
Given the text provided, your task is to transform it into a well-structured document format. Carefully assess the initial sentences to grasp the tone, writing style, and overarching format. Your primary objective is to extend these attributes throughout the entire document, ensuring that it maintains a consistent tone, style, and format as though it was crafted by a single author.
Focus on the following aspects to create a coherent and unified document:
- Core Elements Identification: Determine the main theme to establish an apt title and use any mentioned topics or keywords as headings or subheadings to logically organize the content.
- Uniform Tone and Style: Extend the initial tone and style across the document, ensuring smooth transitions between sections and preserving the text's essence and personality.
- Coherence and Consistency: Pay close attention to sentence structure, word choice, and overall coherence to ensure the document feels naturally consistent.
- Document Structure: Ensure the text is well-organized and properly aligned, maintaining the original flow, tone, and writing style, adjusting only for clarity and coherence.
Final Output: Your reformatted document should be presented in Markdown format. This directive aims to enhance readability and clarity, adhering to the document's structured requirements. The essence of the original text should be evident, with modifications serving to improve the organization and presentation.
Do not include these instructions in your output.
Instructions End."""

    content_part = f"Content to Process:\n{text}\nContent End."

    full_prompt = f"{instructional_part}\n\n{content_part}"

    response = openai.chat.completions.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are a document conversion system. Your task is to convert the provided text into a uniform formatted document while preserving the tone and writing style consistent with the original content."},
            {"role": "user", "content": full_prompt}
        ],
        max_tokens=4000,
        temperature=0.7
    )

    second_extracted_elements = response.choices[0].message.content
    # print(extracted_elements)

    if isinstance(second_extracted_elements, list):
        second_extracted_elements = " ".join(second_extracted_elements)

    return second_extracted_elements


# Function to convert processed Markdown back to DOCX
def markdown_to_docx(markdown_content, output_docx_path):
    pandoc_path = os.getenv('PANDOC_PATH')
    if pandoc_path is None:
        raise ValueError("PANDOC_PATH environment variable is not set.")

    with open("final.md", "w", encoding="utf-8") as md_file:
        md_file.write(markdown_content)

    try:
        subprocess.run([pandoc_path, "-s", "final.md", "-o", output_docx_path], check=True)
    except subprocess.CalledProcessError as e:
        raise ValueError(f"Error during DOCX conversion: {e}")
    finally:
        os.remove("final.md")

# Main function to execute the workflow
def main(docx_file_path, output_docx_path, media_path):
    # Step 1: Convert DOCX to Markdown and extract media
    md_path = convert_docx_to_md_and_extract_media(docx_file_path, media_path)
    
    # Step 2: Read and process the Markdown content for style uniformity
    with open(md_path, "r", encoding="utf-8") as md_file:
        markdown_content = md_file.read()
    uniform_markdown = redefine_from_markdown(markdown_content, style_prompt)
    
    # Step 3: Process the uniform Markdown for structured document formatting
    structured_content = process_text_with_prompt(uniform_markdown)
    
    # Step 4: Convert the final structured content back to a DOCX document
    markdown_to_docx(structured_content, output_docx_path)
    print("Conversion completed successfully.")


