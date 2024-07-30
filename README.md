# PPTX Translator using OpenAI API

PPTX Translator is a Python script to translate text in PowerPoint presentations to a specified language using OpenAI's API

The goal of this is to **retain original formatting** of the powerpoint and only translate the text

The benefits of using OpenAI for translations versus google translate is we do not need to worry about typos in user input for target language, and we do not need to follow the ISO code for languages.

The cons are that I have noticed a increase in time to complete translations.

<add demo video here>




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
