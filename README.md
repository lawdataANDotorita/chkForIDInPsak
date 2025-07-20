# Psak Data Processor

This Python script processes JSON data from the lawdata.co.il API to find cases containing 8-10 digit strings in the text field.

## Requirements

- Python 3.6 or higher
- requests library

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the script:
```bash
python process_psak_data.py
```

## What it does

1. Fetches JSON data from `https://www.lawdata.co.il/lawdata_face_lift_test/chkForIDInPsak.asp`
2. Parses the JSON and extracts the 'data' array
3. Iterates through each object in the array (containing 'c', 'tik', and 'text' fields)
4. Searches for text containing 8-10 digit strings using regex pattern `\b\d{8,10}\b`
5. Writes matching results to `results.txt` with tab-delimited format: `c_value\ttik_value`

## Output

The script creates a `results.txt` file containing the 'c' and 'tik' values for each match, separated by tabs. 