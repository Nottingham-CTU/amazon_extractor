Application to extract and collate amazon voucher URLs in a CSV file, from individual emails exported from Outlook. Supports old and new Outlook for Windows, and new Outlook for Mac.

# User instructions
1. Download the latest version of extract_links.exe from the releases page: https://github.com/Nottingham-CTU/amazon_extractor/releases
2. Create a folder on your PC and drag the voucher emails from Outlook into the folder
3. Open extract_links.exe
4. Click ‘Select folder’
5. Navigate to the folder you created earlier, which contains the emails
6. Click ‘Select folder’
7. Follow the instructions in the dialog prompts
8. A CSV file will be created in the same folder as the emails

# Developer instructions
## Preqs
Make sure you have python 3 installed

## Install dependencies
1. cd into the repo directory if not there already
2. Create a python virtual environment to manage dependences:
``` 
python -m venv extractor_env

```
3. Activate the virtual environment:
```
extractor_env\Scripts\activate
```
4. Install the dependencies with pip:
``` 
pip install beautifulsoup4 lxml extract_msg pyinstaller
```


## Run
To run on your development machine:
```
python extract_links.py
```

## Bundle
To bundle as a Windows .exe for TM team to use:
```
pyinstaller --onefile --noconsole extract_links.py
```
