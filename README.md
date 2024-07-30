# PPTX Translator using OpenAI API

PPTX Translator is a Python script to translate text in PowerPoint presentations to a specified language using OpenAI's API

The goal of this is to **retain original formatting** of the powerpoint and only translate the text



https://github.com/user-attachments/assets/05b50d1f-ac38-4c69-a82d-4d7c326cd904




## Requirements

You need an API key from OpenAI. You can get one from [https://openai.com/index/openai-api/](https://openai.com/index/openai-api/)
- Go to 'Dashboard'
- Click on 'API Keys'
- Create a new secret key
- Create a .env file at same location as script and add 'OPENAI_API_KEY=YOUR_KEY_HERE'

## Usage
```console
python3 translatePPTX.py <PPTX file you want to translate or folder with files> <target language>
```

Example Usage 
```console
python3 translatePPTX.py myPowerpoint.pptx German
```

## Packages

You need to install the following Python packages:

```sh
pip install requests python-pptx tqdm
pip install python-dotenv
```
