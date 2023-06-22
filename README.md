# ofcalX
Introducing a Python application that streamlines the conversion of Google Forms responses into DOCX format. This efficient tool takes each entry from the Google Forms response datasheet and generates separate case documents. Simplify your workflow and save time with this user-friendly solution.

# Setup Guide

Follow the steps below to set up and run the code for converting Google Forms responses to DOCX format:
# Prerequisites

  1. Python 3.x installed on your machine
  2. Required packages: openpyxl, python-docx

# Installation

  1. Clone or download the repository to your local machine.
  2. Open a terminal or command prompt and navigate to the project directory.
  3. Install the required packages: pip install openpyxl python-docx

# Usage

  1. Place your input Excel file (input.xlsx) in the project directory. Make sure it follows the expected format.

  2. Create or update the Word document template (template.docx) according to your desired layout and placeholders. Make sure to enclose placeholders within double curly braces ({{placeholder}}).

  3. Open the Python script file (ofcalExport.py) and make sure to set the correct filenames for the input Excel file and the Word document template.

  4. Run the script: python ofcalExport.py

  5. The resulting DOCX files will be saved in the exports directory. Each generated file will be named based on the first character of the corresponding data row's values.

  6. Check the terminal/console for a success message confirming the generation of Word documents.
