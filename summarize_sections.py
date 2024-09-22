import os
from pptx import Presentation
from extract_sections import extract_sections_and_images
from mistral_summarizer import mistral_summarize
from pptx_exp import create_presentation
import spacy
import re 

# Load spaCy model
nlp = spacy.load('en_core_web_lg')

def top_sentences(text):
    summarized_text = ""
    try:
        # Clean up the text before processing
        text = re.sub(r'\[\d+]+' , '', text)
        text = text.replace("\n", " ")

        # Process text using spaCy
        doc = nlp(text)
        
        # Create a list of (sentence, score) tuples based on sentence similarity
        sentences = [(sent.text.strip(), sent.similarity(doc)) for sent in doc.sents]

        # Sort sentences by similarity and pick the top 5
        top_sentences = sorted(sentences, key=lambda x: x[1], reverse=True)[:5]

        # Combine the top sentences into the final summarized text
        for sentence, score in top_sentences:
            summarized_text += sentence + " "

    except Exception as e:
        print(f"Error in summarizing text: {e}")

    return summarized_text

def summarize_sections(content_dict):
    """
    Summarizes each section from the content_dict using spaCy.
    """
    summarized_dict = {}
    
    for section, content in content_dict.items():
        text = content.get("text", "")
        summarized_text = top_sentences(text)
        summarized_dict[section] = summarized_text
        print(f"Summarized text for section '{section}': {summarized_text}")

    return summarized_dict

def send_to_mistral_for_bullet_points(summarized_dict):
    """
    Sends each section (already summarized using spaCy) to Mistral for bullet points generation.
    """
    bullet_point_dict = {}

    for section, summary in summarized_dict.items():
        if summary:
            bullet_points = mistral_summarize(summary)
            if bullet_points:
                bullet_point_dict[section] = bullet_points
            else:
                bullet_point_dict[section] = "No bullet points available"
        print(f"Bullet points for section '{section}': {bullet_point_dict[section]}")
    
    return bullet_point_dict

if __name__ == "__main__":
    try:
        pdf_filename = 'Introduction to Module_ Neural Networks-1-3.pdf'

        # Step 1: Extract sections and images from the PDF
        content_dict = extract_sections_and_images(pdf_filename)
        print(f"Extracted content: {content_dict}")

        # Step 2: Summarize each section using spaCy
        summarized_dict = summarize_sections(content_dict)

        # Step 3: Send summarized sections to Mistral to get bullet points
        bullet_point_dict = send_to_mistral_for_bullet_points(summarized_dict)

        # Step 4: Prepare images dictionary
        images_dict = {section: content.get("images", []) for section, content in content_dict.items()}

        # Step 5: Choose the PowerPoint template
        folder_path = r'D:/python/Projects/slide_generation/last_hope/templates'
        pptx_files = [f for f in os.listdir(folder_path) if f.endswith('.pptx')]

        print("Choose a PPTX template file to load:")
        for i, file in enumerate(pptx_files):
            print(f"{i}: {file}")

        # Let the user choose a template file
        file_choice = int(input("Enter the number of the file you want to edit: "))
        ppt_template_path = os.path.join(folder_path, pptx_files[file_choice])

        # Step 6: Choose the font
        font_choices = ["Arial", "Calibri", "Times New Roman", "Verdana", "Georgia"]
        print("\nChoose a font:")
        for i, font in enumerate(font_choices):
            print(f"{i}: {font}")

        font_choice = int(input("Enter the number of the font you want to use: "))
        selected_font = font_choices[font_choice]

        # Step 7: Load the presentation template
        print(f"Loading presentation from template: {ppt_template_path}")
        prs = Presentation(ppt_template_path)

        # Step 8: Generate slides using the bullet points and images
        create_presentation(prs, bullet_point_dict, images_dict, selected_font)

        # Step 9: Save the generated presentation
        output_ppt_filename = 'generated_presentation.pptx'
        prs.save(output_ppt_filename)
        print(f"Presentation saved as {output_ppt_filename}")
        
    except Exception as e:
        print(f"An error occurred: {e}")
