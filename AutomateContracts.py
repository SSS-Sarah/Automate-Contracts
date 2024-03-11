# pip install python-docx
from docx import Document
from tkinter import Tk, Label, Entry, Button, messagebox

# creates a custom contract based on input
def create_contract(parent_name, student_name, rate_per_session, num_of_subjects, subject_names):
    template_doc = Document("HLL Parent-Tutor Agreement Template.docx")  # template file name
    print("Template found")
    
    # Find and replace placeholder text in paragraphs
    for paragraph in template_doc.paragraphs:
        paragraph.text = paragraph.text.replace("[PARENT_NAME]", parent_name)
        paragraph.text = paragraph.text.replace("[STUDENT_NAME]", student_name)
        paragraph.text = paragraph.text.replace("[PAYMENT_RATE]", str(rate_per_session))

        # Replace subject placeholders with a loop
        if "[SUBJECTS]" in paragraph.text:
            subject_list = []
            for i, subject_name in enumerate(subject_names):
                subject_list.append(f"\n  * {subject_name}")
            paragraph.text = paragraph.text.replace("[SUBJECTS]", ''.join(subject_list))

    # Save the document in the same directory
    template_doc.save(f"{student_name}_contract.docx") # script ends here
    print("Contract created successfully.")


# Gets information from the GUI and calls the create_contract function
def get_info_and_create_contract():
    
    parent_name = parent_name_entry.get()
    student_name = student_name_entry.get()
    rate_per_session = float(rate_entry.get())
    num_of_subjects = int(num_of_subjects_entry.get())
    
    subject_names = []
    for i in range(0, num_of_subjects):
        subject_names.append(subject_entries[i].get())

    create_contract(parent_name, student_name, rate_per_session, num_of_subjects, subject_names)

    # Clear the entry fields after creating the contract
    parent_name_entry.delete(0, 'end')
    student_name_entry.delete(0, 'end')
    rate_entry.delete(0, 'end')
    num_of_subjects_entry.delete(0, 'end')
    for subject_entry in subject_entries:
        subject_entry.delete(0, 'end')

# Creates subject fields for user input
def create_subject_fields():
    
    global subject_num # global declaration to access it inside function
    
    # if subjectNum is changed, this clears existing fields and entries list.
    for subject_entry in subject_entries:
        subject_entry.destroy()
    subject_entries.clear()
    
    subject_num = int(num_of_subjects_entry.get() or 0) # rereading
    
    for i in range(1, subject_num + 1):
        subject_label = Label(window, text=f"Subject {i}:")
        subject_label.grid(row=6 + i, column=0)
        subject_entry = Entry(window)
        subject_entry.grid(row=6 + i, column=1)
        subject_entries.append(subject_entry)

# Create the GUI elements
window = Tk()
window.title("Tutoring Contract Generator")

# Parent_name input
parent_name_label = Label(window, text="Parent name:")
parent_name_label.grid(row=0, column=0)
parent_name_entry = Entry(window)
parent_name_entry.grid(row=0, column=1)

# Student name input
student_name_label = Label(window, text="Student name:")
student_name_label.grid(row=1, column=0)
student_name_entry = Entry(window)
student_name_entry.grid(row=1, column=1)

# Rate per session input
rate_label = Label(window, text="Rate per session:")
rate_label.grid(row=2, column=0)
rate_entry = Entry(window)
rate_entry.grid(row=2, column=1)

# Number of subjects input
num_of_subjects_label = Label(window, text="Number of subjects:")
num_of_subjects_label.grid(row=3, column=0)
num_of_subjects_entry = Entry(window)
num_of_subjects_entry.grid(row=3, column=1)

subject_num = int(num_of_subjects_entry.get() or 0) # if empty string, sets 0 as default
subject_entries = [] # list to store the names of subjects

# Trigger the creation of subject fields
subject_fields_button = Button(window, text="Add Subject names below", command=create_subject_fields)
subject_fields_button.grid(row=5, column=0)

# Create generate button
create_contract_button = Button(window, text="Create Contract", command=get_info_and_create_contract)
create_contract_button.grid(row= 50 + subject_num, columnspan=2)

window.mainloop()