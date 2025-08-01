import requests
import json
import re
import os
import sys
import glob
import win32com.client
import atexit
import unicodedata

basePath = r'd:\inetpub\wwwroot\upload\psakdin\\'
#basePath = r'c:\users\shay\alltmp\tmppsak\\'


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

def clean_json(input_text):
    """
    Clean response text to remove problematic characters that can cause JSON parsing to fail.
    This function removes:
    - Unicode control characters (except common ones like \n, \r, \t)
    - Zero-width characters
    - Other problematic Unicode characters
    - BOM (Byte Order Mark) characters
    - Non-breaking spaces
    """
    if not input_text:
        return input_text
    
    # Remove BOM characters
    cleaned_text = input_text.replace('\ufeff', '').replace('\ufffe', '')
    
    # Remove zero-width characters
    cleaned_text = re.sub(r'[\u200B-\u200D\uFEFF]', '', cleaned_text)
    
    # Remove problematic Unicode control characters (except \n, \r, \t)
    # This regex removes control characters but preserves \n, \r, \t
    cleaned_text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F\x80-\x9F]', '', cleaned_text)
    
    # Remove non-breaking spaces
    cleaned_text = re.sub(r'\u00A0', ' ', cleaned_text)
    
    # Remove null bytes
    cleaned_text = cleaned_text.replace('\x00', '')
    
    return cleaned_text

def fetch_json_data(url,currentC):
    """Fetch JSON data from the specified URL"""
    try:
        response = requests.get(url,params={"c":currentC})
        response.raise_for_status()  # Raise an exception for bad status codes
        
        # Get the response text and clean it before parsing as JSON
        response_text = response.text
        cleaned_text = clean_json(response_text)
        # Try to parse the cleaned text as JSON
        return json.loads(cleaned_text)
    except requests.RequestException as e:
        print(f"Error fetching data from URL: {e}")
        input("Press any key to continue...")
        return None
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}")
        input("Press any key to continue...")
        return None

def find_digit_strings(text):
    """Find strings composed only of digits with length 8-10 characters, excluding those after 'תיק חיצוני'"""
    if not text:
        return []
    
    # First, find all 8-10 digit strings
    all_digits_pattern = r'(?<![\d_])(\d{7,9}\s*-\s*\d|\d{7,10})(?![\d_])'
    all_matches = re.findall(all_digits_pattern, str(text))
    
    # Then, find digits that are followed by 'תיק חיצוני' (using lookahead)
    excluded_pattern = r'(תיק\s*חיצוני[:\s,\.!?-]*)(\d{6,})'
    excluded_matches = [match[1] for match in re.findall(excluded_pattern, str(text))]

    # Exclude phone numbers, both mobile and landline
    excluded_entities_pattern1 = r'(\d{2,}\s*-\s*)(\d{7,10})'
    excluded_entities_matches1 = [match[1] for match in re.findall(excluded_entities_pattern1, str(text))]

    excluded_entities_pattern2 = r'(\d{7,10})(\s*-\s*[\d]{2,})'
    excluded_entities_matches2 = [match[0] for match in re.findall(excluded_entities_pattern2, str(text))]

    excluded_matches += excluded_entities_matches1 + excluded_entities_matches2

    # Return only matches that are not in the excluded set
    return [match for match in all_matches if match not in excluded_matches]

def check_single_instance():
    """Check if only one instance is running using mutex.txt file"""
    mutex_file = os.path.join(get_script_dir(), "mutex.txt")
    
    if os.path.exists(mutex_file):
        return False
    else:
        # Create the mutex file
        try:
            with open(mutex_file, 'w') as f:
                f.write(str(os.getpid()))
            print(f"Mutex file created. PID: {os.getpid()}")
            return True
        except Exception as e:
            print(f"Error creating mutex file: {e}")
            return False

def cleanup_mutex():
    """Delete the mutex file when the program exits"""
    mutex_file = os.path.join(get_script_dir(), "mutex.txt")
    try:
        if os.path.exists(mutex_file):
            os.remove(mutex_file)
            print("Mutex file cleaned up")
    except Exception as e:
        print(f"Error cleaning up mutex file: {e}")

def process_psak_data():
    """Main function to process the psak data"""

    # Check for single instance
    if not check_single_instance():
        exit()
    
    # Register cleanup function to run on exit
    atexit.register(cleanup_mutex)

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
        for c_value_in_results, tik_value_in_results in results:
            f.write(f"{c_value_in_results}\t{tik_value_in_results}\n")

    # Update currentC.txt with the last c value
    with open(current_c_file, "w", encoding="utf-8") as f:
        f.write(str(c_value))

if __name__ == "__main__":
    process_psak_data() 