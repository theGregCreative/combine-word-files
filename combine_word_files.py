from docx import Document
from tkinter import Tk, filedialog

def combine_word_documents(docx_files, output_file):
    merged_document = Document(docx_files[0])  # Start with the first document

    for file in docx_files[1:]:
        sub_doc = Document(file)
        for element in sub_doc.element.body:
            merged_document.element.body.append(element)

    merged_document.save(output_file)

if __name__ == "__main__":
    # Hide the root window
    root = Tk()
    root.withdraw()

    # Open file dialog to select files
    files_to_merge = filedialog.askopenfilenames(
        title="Select Word Documents to Merge",
        filetypes=[("Word Documents", "*.docx")]
    )

    if files_to_merge:
        combine_word_documents(files_to_merge, "combined.docx")
        print("Documents merged successfully into combined.docx")
    else:
        print("No files selected.")