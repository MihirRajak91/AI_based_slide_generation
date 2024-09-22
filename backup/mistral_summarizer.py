import requests
import json

def mistral_summarize(content):
    """
    Summarizes the given content using Mistral API.
    This function sends a request to the Mistral API and returns bullet points for presentation.
    
    :param content: The text to be summarized.
    :return: A string with bullet points for presentation.
    """
    # Define the URL of the Mistral API endpoint
    url = "http://127.0.0.1:11434/api/generate"  # Ensure this is the correct API endpoint
    
    # Define the headers for the API request
    headers = {
        'Content-Type': 'application/json',
    }
    
    # Prompt for Mistral
    prompt = (
        "I am giving you a paragraph. Return a topic and summary in bullet points. "
        "Strictly follow the format 'Topic: [topic goes here], Summary: bullet point 1, bullet point 2, bullet point 3, bullet point 4, bullet point 5'. "
        "Make the bullet points concise and presentation-friendly.\n\n"
        f"Content:\n{content}"
    )

    # Define the data to be sent in the API request
    data = {
        "model": "mistral", 
        "prompt": prompt,    
        "temperature": 0.3,  
        "max_tokens": 1000   
    }

    try:
        # Send the POST request to the Mistral API
        response = requests.post(url, json=data, headers=headers)
        response.raise_for_status()  # Raise an error for bad responses

        # Parse the JSON response and return the summarized text
        summarized_text = ""
        lines = response.text.strip().splitlines()

        for line in lines:
            try:
                json_line = json.loads(line)  # Try to parse each line as JSON
                if 'response' in json_line:
                    summarized_text += json_line['response']  # Append the summarized text
            except json.JSONDecodeError as e:
                print(f"Error decoding JSON: {e}, skipping line: {line}")

        if summarized_text.strip():
            return summarized_text.strip()
        else:
            print("No valid response found.")
            return None

    except requests.exceptions.RequestException as e:
        print(f"Error during Mistral API call: {e}")
        return None
