Script to extract and collate amazon voucher URLs in a CSV file, from individual emails exported from Outlook. Support old and new Outlook for Windows and new Outlook for Mac.

# Preqs
Make sure you have python 3 installed

# Install dependencies
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


# Run
To run on your development machine:
```
python extract_links.py
```

# Bundle
To bundle as a Windows .exe for TM team to use:
```
pyinstaller --onefile --noconsole extract_links.py
```