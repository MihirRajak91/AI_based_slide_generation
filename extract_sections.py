import pdfplumber
import re
import os

def extract_sections_and_images(filename, image_output_dir="extracted_images"):
    """
    Extract text and images from the PDF, identify sections based on font size.
    Save sections and their content (text + high-quality images) in a dictionary.
    """
    content_dict = {}  # Dictionary to store sections and their content
    current_section = "Introduction"  # Default section
    font_threshold = None  # Threshold for detecting section headings
    section_buffer = []  # Buffer to collect words for section titles

    # Ensure the output directory for images exists
    if not os.path.exists(image_output_dir):
        os.makedirs(image_output_dir)

    with pdfplumber.open(filename) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"Processing page {page_num + 1}...")

            # Extract text and font size
            words = page.extract_words(extra_attrs=['fontname', 'size'])

            for word in words:
                text = word['text']
                font_size = word['size']

                # Set a threshold for identifying section headers based on font size
                if font_threshold is None:
                    font_threshold = determine_font_threshold(words)

                # Collect words for section titles based on large font size
                if font_size >= font_threshold:
                    section_buffer.append(clean_extracted_text(text))

                else:
                    # If there are collected section words, join them into a title
                    if section_buffer:
                        current_section = ' '.join(section_buffer).strip()
                        section_buffer = []  # Clear the buffer
                        content_dict[current_section] = {"text": "", "images": []}  # Create a new section

                    # Append the text to the current section
                    if current_section not in content_dict:
                        content_dict[current_section] = {"text": "", "images": []}
                    content_dict[current_section]["text"] += f"{text} "

            # Extract high-quality images from the current page
            page_images = page.images
            for img_index, img in enumerate(page_images):
                image_bbox = (img['x0'], img['top'], img['x1'], img['bottom'])
                page_image = page.within_bbox(image_bbox).to_image(resolution=300)

                # Save the image to a file with higher quality
                image_filename = f"{image_output_dir}/page_{page_num+1}_image_{img_index+1}.png"
                page_image.save(image_filename, format="PNG", optimize=True, quality=95)

                # Add image to the current section
                content_dict[current_section]["images"].append(image_filename)

            print(f"Images for section '{current_section}': {content_dict[current_section]['images']}")

    print("Finished extracting sections and images.")
    return content_dict

# Add error handling for determining font threshold
def determine_font_threshold(words):
    try:
        font_sizes = [word['size'] for word in words]
        average_size = sum(font_sizes) / len(font_sizes)
        return average_size * 1.5
    except Exception as e:
        print(f"Error determining font threshold: {e}")
        return None

def clean_extracted_text(text):
    return re.sub(r'\s+', ' ', text).strip()

