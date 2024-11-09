# School Transfer Certificate Generator

## Project Overview
This project is a GUI-based application built using Python and Tkinter to facilitate the generation of transfer certificates (TC) for students. It allows users to input a scholar number, retrieve relevant data from a CSV file (`tc_format.csv`), and generate a transfer certificate in a Word document format. This application uses data from the school's database to fill out the certificate and provides an option to customize the student's details.

## Features
- **Load Data from CSV**: Reads data from `tc_format.csv` and `credentials.csv`, which store student details and additional information required for transfer certificates.
- **Generate and Edit Certificates**: Automatically populates a Word document template with student details based on their scholar number. Users can manually edit specific fields such as the student's name, parent's names, attendance details, and issue date.
- **Save and Open Certificate**: Saves the generated certificate in a specified file and opens it automatically for review.

## File Structure
- **`source_code.py`**: The main Python script containing the application's code.
- **`tc_format.csv`**: A CSV file with the base format and placeholders for student data.
- **`credentials.csv`**: A supplementary CSV file providing additional data for fields referenced in the transfer certificate.
- **`ex.docx`**: The template Word document used to generate the transfer certificate.
- **`.gitignore`**: Specifies files to be ignored by version control.

## Prerequisites
To run this project, make sure you have the following:
- **Python 3**: [Download Python 3](https://www.python.org/downloads/) if you don't have it installed.
- **Required Libraries**:
  - Tkinter: Usually included with Python installations.
  - Python-docx: For handling Word documents.
  - Pillow (PIL): For handling images in Tkinter.

To install the additional libraries, run:
```bash
pip install python-docx pillow
