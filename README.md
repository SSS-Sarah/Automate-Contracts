# Automate Contracts
This Python script automates the creation of tutoring contract agreements. It streamlines the process by generating a contract document based on user-provided information and a template DOCX file.

## Getting Started:
This script requires the python-docx library. Install it from the terminal or command prompt:  

    pip install python-docx

## Prepare the Template:
Ensure you have a DOCX file containing the contract template with placeholders for the following:  
**[PARENT_NAME]  
[STUDENT_NAME]  
[PAYMENT_RATE]  
[SUBJECTS]** (This placeholder will be replaced with a list of subjects)

## Run the Script:
Save the script as a Python file. Ensure that it is placed in the same directory as the template contract docx file. To run from terminal:  
```
python automate_contracts.py
```
## Using the Script:

The script will prompt you to enter the following information:   
  
**Parent Name  
Student Name  
Payment Rate per Session  
Number of Subjects**  
  
For the number of subjects, the script will create entry fields for you to enter **each subject name**. Finally, it will generate a new DOCX file with the completed contract information.

## Note:
I created this as an internal tool for a tutoring startup, **Hossain's Legacy of Learning or HLL**, based out of Canada and Bangladesh. It is intended as a basic tool to automate contract generation and might require adjustments based on your specific contract template and workflow.
It is recommended to review the generated document before finalizing the contract.
