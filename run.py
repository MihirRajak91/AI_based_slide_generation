import spacy
from extract_sections import extract_sections_and_images
from summarize_sections import summarize_sections, send_to_mistral_for_bullet_points
from generate_presentation import generate_slides
from pptx import Presentation
import os

# Load spaCy model for summarization
nlp = spacy.load('en_core_web_lg')

def generate_presentation_from_pdf(pdf_filename, ppt_template_path, output_ppt_filename, selected_font="Calibri"):
    """
    This function takes a PDF file, extracts content, summarizes it, and generates a PowerPoint presentation.

    :param pdf_filename: The path to the input PDF file.
    :param ppt_template_path: The path to the PowerPoint template file (.pptx).
    :param output_ppt_filename: The name of the output PowerPoint presentation file.
    :param selected_font: Font choice for the presentation (default is Calibri).
    """
    # Step 1: Extract sections and images from the PDF
    print(f"Extracting sections and images from {pdf_filename}...")
    content_dict = extract_sections_and_images(pdf_filename)

    # Step 2: Summarize each section using spaCy
    print("Summarizing sections...")
    summarized_dict = summarize_sections(content_dict)

    # Step 3: Send the summarized sections to Mistral to get bullet points
    print("Generating bullet points with Mistral...")
    bullet_point_dict = send_to_mistral_for_bullet_points(summarized_dict)

    # Step 4: Prepare the images dictionary
    print("Preparing images...")
    images_dict = {section: content.get('images', []) for section, content in content_dict.items()}

    # Step 5: Load a presentation template
    if not os.path.exists(ppt_template_path):
        raise FileNotFoundError(f"Template file {ppt_template_path} not found.")
    print(f"Loading PowerPoint template from {ppt_template_path}...")
    prs = Presentation(ppt_template_path)

    # Step 6: Generate slides
    print("Generating slides...")
    generate_slides(prs, bullet_point_dict, images_dict, selected_font)

    # Step 7: Save the generated presentation
    print(f"Saving the presentation to {output_ppt_filename}...")
    prs.save(output_ppt_filename)
    print(f"Presentation saved as {output_ppt_filename}.")

# Example usage
if __name__ == "__main__":
    # Input PDF file
    pdf_filename = input("Enter the path to the PDF file: ")

    # PowerPoint template
    ppt_template_path = input("Enter the path to the PowerPoint template (.pptx): ")

    # Output file name
    output_ppt_filename = input("Enter the desired output filename for the PowerPoint presentation (.pptx): ")

    # Optional font choice
    selected_font = input("Enter the font you want to use for the presentation (default is Calibri): ")
    if not selected_font:
        selected_font = "Calibri"  # Default font

    # Generate the presentation
    generate_presentation_from_pdf(pdf_filename, ppt_template_path, output_ppt_filename, selected_font)
