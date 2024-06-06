# Auto-Fill Google Forms with Excel and Capmonster API

### A simple script written in Python to auto-fill and complete Google Forms

This project automates the process of filling out Google Forms using data from an Excel spreadsheet. Written in Python, the script reads data from an Excel file and inputs it into a specified Google Form. Additionally, the project integrates with the CapMonster API to automatically solve captchas, ensuring seamless and uninterrupted form submissions. This tool is designed to save time and effort, especially for tasks involving repetitive form entries.

### Supported Form Features
The script can automatically fill up Google Forms that includes both single and multi-paged forms and can auto-fill the following input styles:
* Text
* TextArea
* Dropdown
* Radio Group
* Scale Group
* Check Group

## Requirements
The required libraries in the script is found in the requirements.txt file. 

### Config Guide
One must also create a config.ini file containing the following information:

**[web-info]**

**url:** The URL of the google forms where you wish to encode from the spreadsheet

**page-count:**  Number of pages in the google forms (1-n)

**[spreadsheet]**

**spreadsheet-dir:** Directory/location of the spreadsheet you wish to encode

**index_id:** set to "True" if there is an "id" column in your spreadsheet, if not, set to "False" (case-sensitive)

**[capmonster]**

**capmonsterKey:** Your unique API/User key from capmonster

## Running the Script
After setting up the config file in the same directory as the python script, you just have to run the script.

`python spreadsheetTransfer.py`

**IMPORTANT: Make sure that the sequence of columns (from left to right) in the spreadsheet are the same sequence of questions that will be shown in the forms (from top to bottom)**

It is also recommended to kepe the browser window focused as much as possible when running the script.
