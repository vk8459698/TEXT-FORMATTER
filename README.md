<!-- README for Document Text Formatter -->
This Python script automates formatting text in Word documents, applying UK English spelling conventions, expanding abbreviations, formatting initials, and handling name references. It preserves quoted text during the formatting process.
Features

Converts US English spellings to UK English (e.g., "organize" → "organise")
Expands common abbreviations (e.g., "eg" → "for example")
Adds periods to initials (e.g., "J R Smith" → "J. R. Smith")
Shortens subsequent references to people (e.g., "Dr. Jane Smith" → "Dr. Smith" after first mention)
Preserves all quoted text exactly as is
Creates verification files to ensure quotes remain unchanged

# to run:
pip install pywin32
python app.py

# Requirements

Python 3.x
Microsoft Word (for Word document processing)
pywin32 library

# Installation

Ensure you have Python 3.x installed on your system
Install the required dependency:

# bash: pip install pywin32
Usage
There are two main ways to use this tool:
1. Processing Word Documents
pythonfrom formatter import format_text

# Process a Word document
format_text("input.docx", "output.docx")
This will:

Open the input Word document
Apply all formatting rules
Save the formatted text to a new Word document
Create verification files for quoted text

2. Processing Plain Text (without Word)
pythonfrom formatter import process_text_file

# Option 1: Process text from a file
# formatted_text = process_text_file(input_file_path="mytext.txt")

# Option 2: Process text directly
text = """Your text here with "quotes that will be preserved" and words to format."""
# formatted_text = process_text_file(input_text=text)

# Do something with the formatted text
# print(formatted_text)
Command Line Usage
You can run the script directly from command line:
bashpython formatter.py
By default, this will process the sample text included in the script.
Customization
To add or modify the formatting rules:

US to UK Spelling Conversion: Add or modify entries in the us_to_uk dictionary in the apply_rules function.
Abbreviation Expansion: Add additional abbreviation patterns in the regex replacement section of the apply_rules function.

<!-- Verification Files -->
The script creates two files to help verify that quotes are preserved correctly:

original_quotes.txt: Contains all quotes from the original document
preserved_quotes.txt: Contains all quotes after formatting

These files can be compared to ensure that quoted content remains unchanged.
Important Notes

The script requires Microsoft Word to be installed when processing .docx files
Running with administrator privileges may be necessary if Word has security restrictions
For large documents, the processing may take some time
Always back up your original documents before processing

<!-- Troubleshooting -->

File Not Found Error: Ensure you're providing the correct file paths
Word Automation Error: Make sure Microsoft Word is installed and not running in protected mode
Permission Denied: Try running as administrator or check file permissions

<!-- License -->
This code is available under the MIT License.