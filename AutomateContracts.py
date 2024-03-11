# pip install docx
# pip install Document
# tkinter is included as a library in latest python builds

#from docx import Document
from tkinter import Tk, Label, Entry, Button, messagebox

# creates a custom contract based on input
def create_contract(parent_name, student_name, rate_per_session, num_of_subjects, subject_names):
    template_doc = "HLL Parent-Tutor Agreement Template.docx"  # template file name
    print("Found template")

    document = Document(template_doc)

    # Find and replace placeholder text in paragraphs
    for paragraph in document.paragraphs:
        paragraph.text = paragraph.text.replace("[PARENT_NAME]", parent_name)
        paragraph.text = paragraph.text.replace("[STUDENT_NAME]", student_name)
        paragraph.text = paragraph.text.replace("[PAYMENT_RATE]", str(rate_per_session))

        # Replace subject placeholders with a loop
        if "[SUBJECTS]" in paragraph.text:
            subject_list = "" #creates a list to store names of subjects/topics
        for i, subject in enumerate(subject_names):
            subject_list += f"\n  * {subject}"
        paragraph.text = paragraph.text.replace("[SUBJECTS]", subject_list)

    # Save the document in the same directory
    document.save(f"{student_name}_contract.docx") # script ends here
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
rate_label = Label(window, text="Parent name:")
rate_label.grid(row=2, column=0)
rate_entry = Entry(window)
rate_entry.grid(row=2, column=1)

# Number of subjects input
num_of_subjects_label = Label(window, text="Number of subjects:")
num_of_subjects_label.grid(row=3, column=0)
num_of_subjects_entry = Entry(window)
num_of_subjects_entry.grid(row=3, column=1)

# Label for subject name entries
subject_label = Label(window, text="Subject Names:")
subject_label.grid(row=4, column=0)

# Create subjectName entry fields based on num_of_subjects input
subject_entries = [] # list to store the names of subjects
subject_num = int(num_of_subjects_entry.get()) # gets the input for numSubjects

for i in range(1, subject_num + 1):
    subject_label = Label(window, text=f"Subject {i}:")
    subject_label.grid(row=4 + i, column=0)
    subject_entry = Entry(window)
    subject_entry.grid(row=4 + i, column=1)
    subject_entries.append(subject_entry)


# Create button
button = Button(window, text="Create Contract", command=get_info_and_create_contract)
button.grid(row=5, columnspan=2)

window.mainloop()