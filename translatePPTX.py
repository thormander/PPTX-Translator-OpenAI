import os
import argparse
import requests
from pptx import Presentation
from tqdm import tqdm
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

# Get your OpenAI API key from environment variable
API_KEY = os.getenv('OPENAI_API_KEY')

# Check for API key
if not API_KEY:
    raise ValueError("No API key found. Please set the 'OPENAI_API_KEY' environment variable.")

# POST translate text using OpenAI API
def translate_text(text, target_language):
    if not text.strip():
        return text  # Return the text as is if it's empty or only whitespace
    
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }
    
    system_message = f"""You are a translator specializing in PowerPoint presentations. 
    Your task is to translate text to {target_language}. 
    For image attributions or license information, keep proper nouns, abbreviations, and license codes unchanged. 
    Translate only the surrounding text.
    IMPORTANT: Your response must be in the following format:
    [START_TRANSLATION]
    Your translated text here
    [END_TRANSLATION]
    Any explanations or notes should be outside these tags."""

    user_message = f"Translate the following text to {target_language}:\n\n{text}"
    
    body = {
        "model": "gpt-3.5-turbo",
        "messages": [
            {"role": "system", "content": system_message},
            {"role": "user", "content": user_message}
        ],
        "max_tokens": 1000,
        "n": 1,
        "temperature": 0.3
    }
    
    response = requests.post(url, headers=headers, json=body)
    if response.status_code == 200:
        content = response.json()['choices'][0]['message']['content'].strip()
        # Extract only the translated text between the tags
        start_tag = "[START_TRANSLATION]"
        end_tag = "[END_TRANSLATION]"
        start_index = content.find(start_tag) + len(start_tag)
        end_index = content.find(end_tag)
        if start_index != -1 and end_index != -1:
            return content[start_index:end_index].strip()
        else:
            print(f"WARN: Translation tags not found in the response, Skipping. Full response: {content}")
            return text
    else:
        print(f"Error translating text: {response.status_code} {response.text}")
        return text

def translate_shape_text(shape, target_language):
    if not hasattr(shape, "text_frame") or not shape.text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            translated_text = translate_text(run.text, target_language)
            run.text = translated_text

def process_presentation(input_file, target_language):
    print(f"Opening {input_file}")
    try:
        input_ppt = Presentation(input_file)
    except Exception as e:
        print(f"Error opening file {input_file}: {e}")
        return

    slide_count = len(input_ppt.slides)
    
    with tqdm(total=slide_count, desc="Translating", unit="slide") as pbar:
        for i, slide in enumerate(input_ppt.slides):
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame:
                    try:
                        translate_shape_text(shape, target_language)
                    except Exception as e:
                        print(f"Error processing shape on slide {i}: {e}")
            pbar.update(1)

    output_file = f"{target_language}_{os.path.basename(input_file)}"
    try:
        input_ppt.save(output_file)
        print(f"\nSaved as {output_file}")
    except Exception as e:
        print(f"Error saving file {output_file}: {e}")

def process_folder(folder_path, target_language):
    for filename in os.listdir(folder_path):
        if filename.endswith(".pptx"):
            file_path = os.path.join(folder_path, filename)
            process_presentation(file_path, target_language)

def main():
    parser = argparse.ArgumentParser(description="Translate PowerPoint presentations. Usage: python3 translatePPTX.py <input_path> <target_language>")
    parser.add_argument("input_path", nargs='?', help="Path to the input PowerPoint file or folder")
    parser.add_argument("target_language", nargs='?', help="Target language for translation (ex: 'en' for English, 'es' for Spanish)")
    args = parser.parse_args()

    if not args.input_path or not args.target_language:
        parser.print_help()
        return
    
    # handle individual vs bulk handling
    if os.path.isdir(args.input_path):
        process_folder(args.input_path, args.target_language)
    else:
        process_presentation(args.input_path, args.target_language)

if __name__ == "__main__":
    main()
