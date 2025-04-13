import os
import re
import win32com.client

def apply_uk_spelling(text):
    # Dictionary of US to UK spelling replacements
    uk_spelling = {
        'organize': 'organise',
        'organizing': 'organising',
        'organized': 'organised',
        'organizes': 'organises',
    }
    
    # Replace abbreviations
    abbreviations = {
        r'\beg\b': 'for example',
        r'\betc\b': 'etc.',
    }
    
    # Create a pattern for quoted text
    quote_pattern = r'"([^"]*)"'
    quotes = re.findall(quote_pattern, text)
    
    # Replace quotes with placeholders to protect them
    for i, quote in enumerate(quotes):
        text = text.replace(f'"{quote}"', f'__QUOTE{i}__')
    
    # Apply UK spelling changes (avoiding proper nouns)
    for us, uk in uk_spelling.items():
        # Avoid replacing in proper nouns (words with capital first letter)
        text = re.sub(rf'\b(?<![A-Z]){us}\b', uk, text)
    
    # Apply abbreviation replacements (outside quotes)
    for abbr, full in abbreviations.items():
        text = re.sub(abbr, full, text)
    
    # Restore quotes
    for i, quote in enumerate(quotes):
        text = text.replace(f'__QUOTE{i}__', f'"{quote}"')
    
    return text

def format_names(text):
    # Dictionary to track first mentions
    first_mentions = {}
    # Dictionary to track ambiguous last names
    last_name_count = {}
    
    # Process initials
    text = re.sub(r'\b([A-Z])(?= [A-Z][a-z])', r'\1.', text)  # Add periods to initials
    
    # Regular expression for title + name patterns
    name_pattern = r'\b(Dr|Mr|Mrs|Ms|Prof|Professor|Sir|Lady)\s+([A-Z][a-z]+)(?:\s+([A-Z])(?!\.))?(?:\s+)?([A-Z][a-z]+)\b'
    
    # Find names and count last names for ambiguity check
    for match in re.finditer(name_pattern, text):
        title, first, middle, last = match.groups()
        if last in last_name_count:
            last_name_count[last] += 1
        else:
            last_name_count[last] = 1
    
    # Function to replace full names with shortened versions on subsequent mentions
    def replace_names(match):
        title, first, middle, last = match.groups()
        full_name = match.group(0)
        
        # Format initials with periods
        if middle and '.' not in middle:
            full_name_formatted = f"{title} {first} {middle}. {last}"
        else:
            full_name_formatted = full_name
        
        # Check for ambiguity
        if last in last_name_count and last_name_count[last] > 1:
            # If last name is ambiguous, always use full name
            return full_name_formatted
        
        # For non-ambiguous names, use title + last name on subsequent mentions
        key = f"{first}_{last}"
        if key in first_mentions:
            return f"{title} {last}"
        else:
            first_mentions[key] = True
            return full_name_formatted
    
    # Apply the name formatting
    text = re.sub(name_pattern, replace_names, text)
    
    # Handle other initials in names (like Franklin D Roosevelt)
    initial_pattern = r'\b([A-Z][a-z]+)\s+([A-Z])(?!\.)\s+([A-Z][a-z]+)\b'
    text = re.sub(initial_pattern, r'\1 \2. \3', text)
    
    return text

def process_word_document(input_file, output_file):
    # Start Word application
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    
    try:
        # Open the input document
        doc = word_app.Documents.Open(os.path.abspath(input_file))
        
        # Get the text content
        content = ""
        for paragraph in doc.Paragraphs:
            content += paragraph.Range.Text + "\n"
        
        # Apply formatting rules
        content = apply_uk_spelling(content)
        content = format_names(content)
        
        # Create a new document for output
        output_doc = word_app.Documents.Add()
        
        # Add "Corrected:" header
        output_doc.Content.Text = "Corrected:\n" + content
        
        # Save the output document
        output_doc.SaveAs(os.path.abspath(output_file))
        output_doc.Close()
        
        # Close the input document
        doc.Close(False)
        
    finally:
        # Quit Word application
        word_app.Quit()

def main():
    input_file = "input.docx"
    output_file = "output.docx"
    
    print(f"Processing '{input_file}'...")
    process_word_document(input_file, output_file)
    print(f"Formatting complete. Output saved to '{output_file}'")

if __name__ == "__main__":
    main()
