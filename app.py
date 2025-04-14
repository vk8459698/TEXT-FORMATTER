import re
import win32com.client
import os
import uuid

def format_text(input_docx, output_docx):
    # Get absolute paths
    input_path = os.path.abspath(input_docx)
    output_path = os.path.abspath(output_docx)
    
    # Check if input file exists
    if not os.path.exists(input_path):
        print(f"Error: Input file not found at {input_path}")
        return
    
    print(f"Opening file: {input_path}")
    
    # Connect to Word application
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    
    try:
        # Open the document
        doc = word_app.Documents.Open(input_path)
        
        # Get document text
        doc_text = doc.Range().Text
        
        # Save original quoted content to a file (for verification)
        save_quotes_to_file(doc_text, "original_quotes.txt")
        
        # Apply formatting rules
        corrected_text = apply_rules(doc_text)
        
        # Save preserved quoted content to a file (for verification)
        save_quotes_to_file(corrected_text, "preserved_quotes.txt")
        
        # Create new document for output
        output_doc = word_app.Documents.Add()
        output_doc.Range().Text = "Corrected:\n" + corrected_text
        
        # Save and close
        output_doc.SaveAs(output_path)
        output_doc.Close()
        doc.Close(False)  # Don't save changes to original
        
        print(f"Formatting complete. Output saved to: {output_path}")
        print("Quote preservation files created for verification.")
        
    except Exception as e:
        print(f"Error occurred: {str(e)}")
    finally:
        # Always quit Word
        word_app.Quit()

def save_quotes_to_file(text, filename):
    """Extract all quoted text and save to a file for verification"""
    quotes = re.findall(r'"[^"]*"', text)
    with open(filename, 'w', encoding='utf-8') as f:
        for i, quote in enumerate(quotes, 1):
            f.write(f"Quote {i}: {quote}\n")

def apply_rules(text):
    # Step 1: Identify and mark quotes with unique placeholders
    quotes = {}
    quote_pattern = r'"([^"]*)"'
    
    def replace_quote(match):
        quote_content = match.group(0)  # Get the full quoted text including quotes
        placeholder = f"___QUOTE_{uuid.uuid4()}___"
        quotes[placeholder] = quote_content
        return placeholder
    
    # Replace all quotes with placeholders
    text_with_placeholders = re.sub(quote_pattern, replace_quote, text)
    
    # Step 2: Apply formatting rules to the text (excluding quotes)
    # UK spelling
    us_to_uk = { r'\borganize\b': 'organise', r'\borganizes\b': 'organises', r'\borganizing\b': 'organising', r'\borganized\b': 'organised', r'\bcenter\b': 'centre', r'\bcenters\b': 'centres', r'\bcolor\b': 'colour', r'\bcolors\b': 'colours', r'\bfavor\b': 'favour', r'\bfavors\b': 'favours', r'\bfavorite\b': 'favourite', r'\bfavorites\b': 'favourites', r'\blabor\b': 'labour', r'\blabors\b': 'labours', r'\bneighbor\b': 'neighbour', r'\bneighbors\b': 'neighbours', r'\bneighborhood\b': 'neighbourhood', r'\bneighborhoods\b': 'neighbourhoods', r'\brealize\b': 'realise', r'\brealizes\b': 'realises', r'\brealizing\b': 'realising', r'\brealized\b': 'realised', r'\brecognize\b': 'recognise', r'\brecognizes\b': 'recognises', r'\brecognizing\b': 'recognising', r'\brecognized\b': 'recognised', r'\bspecialize\b': 'specialise', r'\bspecializes\b': 'specialises', r'\bspecializing\b': 'specialising', r'\bspecialized\b': 'specialised', r'\bemphasize\b': 'emphasise', r'\bemphasizes\b': 'emphasises', r'\bemphasizing\b': 'emphasising', r'\bemphasized\b': 'emphasised' }
    
    processed_text = text_with_placeholders
    for us_pattern, uk_spelling in us_to_uk.items():
        processed_text = re.sub(us_pattern, uk_spelling, processed_text)
    
    # Replace abbreviations
    processed_text = re.sub(r'\beg\b', 'for example', processed_text)
    
    # Step 3: Add periods to initials
    processed_text = re.sub(r'\b([A-Z])\s([A-Z])\b', r'\1. \2.', processed_text)
    processed_text = re.sub(r'\b([A-Z])\s([A-Z][a-z])', r'\1. \2', processed_text)
    
    # Step 4: Process names for subsequent references
    processed_text = process_names(processed_text)
    
    # Step 5: Restore the original quotes
    for placeholder, quote in quotes.items():
        processed_text = processed_text.replace(placeholder, quote)
    
    return processed_text

def process_names(text):
    """Process names for subsequent references using regex replacements"""
    # First, identify all names with their titles
    name_pattern = r'\b(Dr|Mr|Mrs|Ms|Prof)\.?\s+([A-Z][a-z]+)\s+([A-Z][a-z]+)\b'
    name_matches = re.findall(name_pattern, text)
    
    # Create a dictionary to track unique full names
    names = {}
    for match in name_matches:
        title, first_name, last_name = match
        full_name = f"{title} {first_name} {last_name}"
        if full_name not in names:
            names[full_name] = {
                'title': title,
                'first_name': first_name,
                'last_name': last_name,
                'pattern': re.compile(re.escape(full_name))
            }
    
    # Count shared last names
    last_name_count = {}
    for info in names.values():
        last_name = info['last_name']
        if last_name in last_name_count:
            last_name_count[last_name] += 1
        else:
            last_name_count[last_name] = 1
    
    # Process text for each name
    processed_text = text
    
    for full_name, info in names.items():
        title = info['title']
        last_name = info['last_name']
        pattern = info['pattern']
        
        # Skip if last name is shared with another person
        if last_name_count[last_name] > 1:
            continue
        
        # Count occurrences
        occurrences = pattern.findall(processed_text)
        if len(occurrences) <= 1:
            continue
            
        # Replace all occurrences after the first one
        # First find all matches
        matches = list(pattern.finditer(processed_text))
        if len(matches) <= 1:
            continue
            
        # Start with the last occurrence and work backwards
        # This preserves the positions for earlier replacements
        for match in reversed(matches[1:]):
            start, end = match.span()
            short_name = f"{title} {last_name}"
            before = processed_text[:start]
            after = processed_text[end:]
            processed_text = before + short_name + after
    
    return processed_text

# For testing without Word automation
def process_text_file(input_file_path=None, input_text=None):
    """Process text directly from a file or string without Word document handling"""
    if input_file_path:
        try:
            with open(input_file_path, 'r', encoding='utf-8') as f:
                input_text = f.read()
        except Exception as e:
            print(f"Error reading file: {str(e)}")
            return None
    
    if not input_text:
        print("No input text provided")
        return None
        
    # Process the text
    corrected_text = apply_rules(input_text)
    return corrected_text

# Example usage
if __name__ == "__main__":
 
#     sample_text1 = """Article: A Conference on Global Policy Challenges 

# On Monday, we organize a significant conference eg in London, hosted by the World Health Organization. The event aims to address pressing global policy issues, and Dr Manmohan Singh will deliver the opening remarks. Dr Manmohan Singh has been a key figure in economic policy, and he plans to stay for the entire three-day event. He mentioned, "We organize eg our sessions to maximize engagement," emphasizing the importance of structured discussions, etc. The conference features several notable speakers, including Nawaz Sharif, who will speak on trade policies. Later in the day, Shehbaz Sharif takes the stage to discuss healthcare reforms, creating some confusion among attendees due to their shared last name. Nawaz Sharif focuses on tariffs, while Shehbaz Sharif highlights the need for universal healthcare access. "It's crucial to organize eg partnerships," Shehbaz Sharif told the audience, sharing his vision for collaborative efforts. Another speaker, Franklin D Roosevelt, joins the event as a policy historian. He often references the legacy of Franklin D Roosevelt, the former U.S. president, in his talks, which might puzzle some attendees unfamiliar with the distinction. Franklin D Roosevelt, the historian, prefers to be called "Dr Roosevelt" in formal settings, but the conference program lists him as "franklin d roosevelt." He plans to discuss historical policy impacts, eg during the Great Depression. Dr Aishwarya Rai, a renowned public health expert, also participates in the conference. Dr Aishwarya Rai has collaborated extensively with the World Health Organization on vaccination campaigns, and she frequently uses terms like "eg" in her presentations. During a panel, she stated, "We organize eg our data to ensure clarity." Her session draws a large crowd, eager to hear her insights on global health strategies, etc. The event includes a breakout session led by Dr Singh, who reflects on his time as an economic advisor. Dr Manmohan Singh often cites historical examples to support his arguments, making his session highly engaging. He references the World Health Organization's role in past health crises, noting their efforts to organize eg international responses. Attendees appreciate his depth of knowledge and practical approach. Midway through the conference, Nawaz Sharif and Shehbaz Sharif host a joint Q&A session. The Sharif brothers address questions on economic collaboration, with Nawaz Sharif focusing on trade agreements and Shehbaz Sharif emphasizing healthcare funding. Their shared last name continues to cause minor mix-ups, as some attendees mistakenly attribute comments to the wrong speaker. "We must organize eg better communication," Nawaz Sharif remarked, addressing the confusion. Dr Roosevelt, the historian, presents a detailed analysis of Franklin D Roosevelt's New Deal policies. He contrasts the former president's strategies with modern approaches, using terms like "etc" to summarize lists of reforms. Franklin D Roosevelt, the historian, also shares a quote from a 1930s speech: "We must organize eg for progress." His presentation highlights the relevance of historical lessons in today's policy landscape. In a later session, Dr Rai returns to the stage to discuss her recent work with the World Health Organization. Dr Aishwarya Rai emphasizes the importance of data-driven decision-making, often using shorthand like "eg" in her slides. She mentions, "We organize eg our teams to tackle outbreaks," illustrating her approach to crisis management. Her talk inspires many attendees to consider public health careers. The conference concludes with a closing address by Dr Singh, who summarizes the key takeaways. Dr Manmohan Singh praises the contributions of all speakers, including Nawaz Sharif and Shehbaz Sharif, for their insightful discussions. He also acknowledges Franklin D Roosevelt for his historical perspective, noting how it enriched the event. Dr Singh's speech leaves the audience motivated to apply the lessons learned, etc. Reflecting on the event, Dr Rai shares her final thoughts with the press. Dr Aishwarya Rai highlights the collaborative spirit of the conference, crediting the World Health Organization for its role in bringing experts together. She concludes, "We organize eg these events to foster innovation," underscoring the value of such gatherings. The conference ends on a high note, with attendees eager for next year's edition."""
    


#     corrected1 = process_text_file(input_text=sample_text1)
#     print(corrected1)

          
    format_text("input.docx", "output.docx")