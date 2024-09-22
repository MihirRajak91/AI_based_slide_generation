import requests
import json

def mistral_summarize(content):
    """
    Summarizes the given content using Mistral API.
    This function sends a request to the Mistral API and returns bullet points for presentation.
    """
    url = "http://127.0.0.1:11434/api/generate"
    
    headers = {
        'Content-Type': 'application/json',
    }
    
    prompt = (
        "I am giving you a paragraph. Return a topic and summary in bullet points. "
        "keep the points short and too the point"
        "Strictly follow the format 'Topic: [topic goes here], Summary: bullet point 1, bullet point 2, bullet point 3, bullet point 4, bullet point 5'. "
        "Make the bullet points concise and presentation-friendly.\n\n"
        f"Content:\n{content}"
    )

    data = {
        "model": "mistral", 
        "prompt": prompt,    
        "temperature": 0.3,  
        "max_tokens": 1000   
    }

    try:
        print(f"Sending request to Mistral with prompt: {prompt}")
        response = requests.post(url, json=data, headers=headers)
        response.raise_for_status()

        # Parse the JSON response and return the summarized text
        summarized_text = ""
        lines = response.text.strip().splitlines()

        for line in lines:
            try:
                json_line = json.loads(line)
                if 'response' in json_line:
                    summarized_text += json_line['response']
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
