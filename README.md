# PPTX Translator using OpenAI API

PPTX Translator is a Python script to translate text in PowerPoint presentations to a specified language using OpenAI's API

The goal of this is to **retain original formatting** of the powerpoint and only translate the text

The benefits of using OpenAI for translations versus google translate is we do not need to worry about typos in user input for target language, and we do not need to follow the ISO code for languages.

The cons are that I have noticed a increase in time to complete translations.


https://github.com/user-attachments/assets/e83a68f1-6dc2-4828-9f86-5f0c58f20ce5



## FYI
Currently, I have hardcoded the rate limits in the code. You will need to adjust this based on your use case. For me, I have 'Tier 1' which allows me a max of 3500 RPM and 200000 TPM. As I do not have any other requests or tokens from anywhere else, I have just put the max for Tier 1 into the variables. If you expect your api to be used in other programs, a safe starting point would be 50% of the max allowed per minute. If your endpoint is already being heavily utilized, you may want to decrease the max allowed further to not cause issues for other programs that may hit the max limit.

Rate Limits can be found here: [https://platform.openai.com/docs/guides/rate-limits](https://platform.openai.com/docs/guides/rate-limits)

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
