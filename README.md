# IT23177628 - IT3040 ITPM Assignment 1

## Assignment Option

Option 1: Transliteration Accuracy Testing for the Chat Sinhala transliteration function.

## Student Registration Number

IT23177628

## Application Under Test

https://www.pixelssuite.com/chat-translator

## Scope of Testing

This project tests the Chat Sinhala transliteration function by entering chat-style Singlish inputs into the live application and comparing the generated Sinhala output with the expected Sinhala output recorded in the Excel sheet.

The testing focuses only on transliteration accuracy. Backend API testing, performance testing, scalability testing, and security testing are outside the scope of this assignment.

## Project Structure

```text
IT23177628/
├── README.md
├── Git_Repository_Link_IT23177628.txt
├── requirements.txt
└── test_automation/
    ├── IT23177628_test_automation.py
    └── Assignment 1 - Test cases.xlsx
```

## Required Software

- Python 3
- pip
- Playwright
- openpyxl
- Google Chrome or Playwright Chromium

## Installation Steps for Mac

Open Terminal from the project root folder and run:

```bash
python3 -m pip install -U pip
python3 -m pip install -r requirements.txt
python3 -m playwright install
```

If `requirements.txt` is not used, the dependencies can also be installed manually:

```bash
python3 -m pip install playwright openpyxl
python3 -m playwright install
```

## How to Run the Automated Test

First, go to the project folder:

```bash
cd ~/Desktop/IT23177628
```

Then run the automation command:

```bash
python3 test_automation/IT23177628_test_automation.py --excel "test_automation/Assignment 1 - Test cases.xlsx" --sheet "Test cases" --url "https://www.pixelssuite.com/chat-translator" --wait-ms 10000 --retries 15 --retry-wait-ms 2000 --type-delay-ms 120 --slow-mo-ms 300 --save-every 1 --keep-open
```

## How the Script Works

1. Opens the Excel file `Assignment 1 - Test cases.xlsx`.
2. Reads each Singlish input from the `Input` column.
3. Opens the Chat Translator website.
4. Enters each input into the translator.
5. Captures the generated Sinhala output.
6. Writes the generated output into the `Actual output` column.
7. Compares the actual output with the expected output.
8. Records the test result in the `Status` column.

## Excel Sheet Details

The Excel sheet contains the following columns:

- TC ID
- Input length type
- Input
- Expected output
- Actual output
- Status
- Singlish input types covered
- Evidence or rationale for the input type covered

The test cases are negative test cases and use TC IDs from `Neg_0001` to `Neg_0050`.

## Notes

- Close the Excel file before running the automation script.
- If the browser remains open after execution, press `Control + C` in Terminal to stop the script.
- The public GitHub repository link is included in `Git_Repository_Link_IT23177628.txt`.
