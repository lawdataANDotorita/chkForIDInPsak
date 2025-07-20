import requests
import json
import re
import os
import sys

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
    all_digits_pattern = r'\d{8,10}'
    all_matches = re.findall(all_digits_pattern, str(text))
    
    # Then, find digits that are followed by 'תיק חיצוני' (using lookahead)
    # We'll use multiple patterns to handle different special characters

    excluded_pattern = r'(תיק\s*חיצוני[:\s,\.!?-]*)(\d{8,10})'
    
    # We want excluded_matches to contain only the digits group (group 2 from the regex)
    excluded_matches = [match[1] for match in re.findall(excluded_pattern, str(text))]


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
    url = "https://www.lawdata.co.il/lawdata_face_lift_test/chkForIDInPsak.asp"
    
    # Fetch the JSON data
    print("Fetching data from URL...")


    current_c_file = os.path.join(get_script_dir(), "currentC.txt")
    currentC=0
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
    
    # Write results to file
    output_file = os.path.join(get_script_dir(), "filesWithID.txt")
    with open(output_file, 'a', encoding='utf-8') as f:
        for idx, (c_value, tik_value) in enumerate(results, 1):
            f.write(f"{idx}\t{c_value}\t{tik_value}\n")
    print(f"Processing complete. Found {len(results)} matches.")
    print(f"Results written to {output_file}")

    # Update currentC.txt with the last c value
    with open(current_c_file, "w", encoding="utf-8") as f:
        f.write(str(c_value))

if __name__ == "__main__":
    process_psak_data() 