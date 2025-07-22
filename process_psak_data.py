import requests
import json
import re
import os
import sys
import glob
import win32com.client

basePath = r'd:\inetpub\wwwroot\upload\psakdin\\'
# basePath = r'c:\users\shay\alltmp\\'


def cover_id_in_word_file(c_value,digit_strings):
    # Prepare possible file suffixes
    suffixes = [".docx",".doc",".rtf"]
    # Build glob patterns for each suffix
    file_patterns = [os.path.join(basePath, f"{c_value}{suffix}") for suffix in suffixes]
    # Find all matching files
    matching_files = []
    for pattern in file_patterns:
        matching_files.extend(glob.glob(pattern))

    if not matching_files:
        return

    # Create Word application once for all files
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    try:
        for file_path in matching_files:
            try:
                # Open the document
                doc = word.Documents.Open(file_path)
                
                # First, let's check what text is actually in the document
                doc_text = doc.Content.Text
                print(f"Document contains {len(doc_text)} characters")
                print(f"Looking for digit strings: {digit_strings}")
                
                for digit_str in digit_strings:
                    # Check if the digit string exists in the document
                    if digit_str in doc_text:

                        start = 0
                        doc_range = doc.Content
                        while True:
                            found = doc_range.Find.Execute(FindText=digit_str, Replace=0)
                            if not found:
                                break
                            # Select the found text and replace it
                            doc_range.Text = 'xxxxxxxx'
                            # Move start to after the replaced text
                            start = doc_range.End
                            if start >= doc.Content.End:
                                break
                            doc_range = doc.Range(start, doc.Content.End)
                    else:
                        print(f"Digit string '{digit_str}' NOT found in document")
                
                
                # Save the document - use Save() instead of SaveAs() to preserve original format
                doc.Save()
                doc.Close()
                print(f"Updated Word file: {file_path}")
                
            except Exception as e:
                print(f"Error processing Word file {file_path}: {e}")
                try:
                    doc.Close(SaveChanges=False)
                except:
                    pass
    finally:
        # Always quit Word application
        try:
            word.Quit()
        except:
            pass


def cover_id_in_file(c_value,digit_strings):
    # Prepare possible file suffixes
    suffixes = [".html",".htm",".txt"]
    # Build glob patterns for each suffix
    file_patterns = [os.path.join(basePath, f"{c_value}{suffix}") for suffix in suffixes]
    # Find all matching files
    matching_files = []
    for pattern in file_patterns:
        matching_files.extend(glob.glob(pattern))

    for file_path in matching_files:
        try:
            with open(file_path, "r", encoding="windows-1255") as f:
                file_text = f.read()
            # Replace each digit string with 'xxxxxxxx'
            for digit_str in digit_strings:
                file_text = file_text.replace(digit_str, 'xxxxxxxx')
            with open(file_path, "w", encoding="windows-1255") as f:
                f.write(file_text)
            print(f"Updated file: {file_path}")
        except Exception as e:
            print(f"Error processing file {file_path}: {e}")

def get_script_dir():
    # Get the directory where the script/exe is located
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle (exe)
        return os.path.dirname(sys.executable)
    else:
        # If the application is run from a Python interpreter
        return os.path.dirname(os.path.abspath(__file__))

def fetch_json_data(url,currentC):
    """Fetch JSON data from the specified URL"""
    try:
        response = requests.get(url,params={"c":currentC})
        response.raise_for_status()  # Raise an exception for bad status codes
        return response.json()
    except requests.RequestException as e:
        print(f"Error fetching data from URL: {e}")
        return None
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")
        return None

def find_digit_strings(text):
    """Find strings composed only of digits with length 8-10 characters, excluding those after 'תיק חיצוני'"""
    if not text:
        return []
    
    # First, find all 8-10 digit strings
    all_digits_pattern = r'(?<!\d)(\d{7,9}|\d{7,8}-\d)(?!\d)'
    all_matches = re.findall(all_digits_pattern, str(text))
    
    # Then, find digits that are followed by 'תיק חיצוני' (using lookahead)
    # We'll use multiple patterns to handle different special characters

    excluded_pattern = r'(תיק\s*חיצוני[:\s,\.!?-]*)(\d{6,})'
    excluded_matches = [match[1] for match in re.findall(excluded_pattern, str(text))]

    excluded_phone_pattern = r'(0[0-9]{1,2}\s*-\s*)(\d{7,8})'
    excluded_phone_matches = [match[1] for match in re.findall(excluded_phone_pattern, str(text))]

    excluded_matches += excluded_phone_matches

    if len (all_matches) > 0:
        print(f"Found {len(all_matches)} digit strings")
        print(f"All matches: {all_matches}")
    if len (excluded_matches) > 0:
        print(f"Found {len(excluded_matches)} excluded matches")
        print(f"Excluded matches: {excluded_matches}")


    # Return only matches that are not in the excluded set
    return [match for match in all_matches if match not in excluded_matches]

def process_psak_data():
    """Main function to process the psak data"""
    url = "https://www.lawdata.co.il/chkForIDInPsak.asp"
    
    # Fetch the JSON data
    print("Fetching data from URL...")

    current_c_file = os.path.join(get_script_dir(), "currentC.txt")
    currentC=5000000
    if os.path.exists(current_c_file):
        with open(current_c_file, "r", encoding="utf-8") as f:
            currentC = f.read().strip()
    
    json_data = fetch_json_data(url,currentC)

    if not json_data:
        print("Failed to fetch or parse JSON data")
        return
    
    # Check if 'data' member exists
    if 'data' not in json_data:
        print("No 'data' member found in JSON")
        return
    
    data_array = json_data['data']
    print(f"Found {len(data_array)} items in data array")
    

    # Process each item in the array
    results = []
    for item in data_array:
        # Check if required fields exist
        if 'c' not in item or 'tik' not in item or 'text' not in item:
            continue
        
        c_value = item['c']
        tik_value = item['tik']
        text_value = item['text']
        
        text_value = text_value.replace('\r\n', ' ')

        
        # Find digit strings in the text
        digit_strings = find_digit_strings(text_value)
        
        # If we found any 8-10 digit strings, add to results
        if digit_strings:
            results.append((c_value, tik_value))
            print(f"Found match: c={c_value}, tik={tik_value}, digits={digit_strings}")

            cover_id_in_file(c_value,digit_strings)
            cover_id_in_word_file(c_value,digit_strings)

    # Write results to file
    output_file = os.path.join(get_script_dir(), "filesWithID.txt")
    with open(output_file, 'a', encoding='utf-8') as f:
        for c_value, tik_value in results:
            f.write(f"{c_value}\t{tik_value}\n")

    # Update currentC.txt with the last c value
    with open(current_c_file, "w", encoding="utf-8") as f:
        f.write(str(c_value))

if __name__ == "__main__":
    process_psak_data() 