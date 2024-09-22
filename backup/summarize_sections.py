import re
import spacy
from extract_sections import extract_sections_and_images
from mistral_summarizer import mistral_summarize  # Use the correct import
from generate_presentation import generate_slides  # Import the slide generation function
from pptx import Presentation

# Load spaCy model
nlp = spacy.load('en_core_web_lg')

def top_sentences(text):
    """
    Summarizes the given text using the top 5 sentences based on similarity.
    """
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
        print(f"Error in summarizing: {e}")

    return summarized_text

def summarize_sections(content_dict):
    """
    Summarizes each section from the content_dict using spaCy.
    Returns a dictionary where the title is the key and the summarized section is the value.
    """
    summarized_dict = {}
    
    for section, content in content_dict.items():
        text = content.get("text", "")
        summarized_text = top_sentences(text)
        summarized_dict[section] = summarized_text 

    return summarized_dict

def send_to_mistral_for_bullet_points(summarized_dict):
    """
    Sends each section (already summarized using spaCy) to Mistral for bullet points generation.
    Returns a new dictionary with section titles as keys and Mistral-generated bullet points as values.
    """
    bullet_point_dict = {}

    for section, summary in summarized_dict.items():
        if summary:
            # Send the spaCy-generated summary to Mistral and get bullet points
            bullet_points = mistral_summarize(summary)
            
            if bullet_points:
                bullet_point_dict[section] = bullet_points
            else:
                bullet_point_dict[section] = "No bullet points available"
    
    return bullet_point_dict

if __name__ == "__main__":
    pdf_filename = 'Introduction to Module_ Neural Networks-1-3.pdf'  # Path to the PDF file
    ppt_template_path = 'template.pptx'  # Path to your PowerPoint template
    output_ppt_filename = 'generated_presentation.pptx'  # Name of the output PowerPoint

    # Step 1: Extract sections and images
    content_dict = extract_sections_and_images(pdf_filename)

    # Step 2: Summarize each section using spaCy
    summarized_dict = summarize_sections(content_dict)

    # Step 3: Send the summarized sections to Mistral to get bullet points
    bullet_point_dict = send_to_mistral_for_bullet_points(summarized_dict)

    # Step 4: Prepare the images dictionary
    images_dict = {section: content.get("images", []) for section, content in content_dict.items()}

    # Step 5: Load a presentation template
    prs = Presentation(ppt_template_path)

    # Step 6: Generate slides using the bullet points and images
    selected_font = "Calibri"  # You can allow user input for this
    generate_slides(prs, bullet_point_dict, images_dict, selected_font)

    # Step 7: Save the updated presentation
    prs.save(output_ppt_filename)
    print(f"Presentation saved as {output_ppt_filename}.")
