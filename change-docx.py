import os
import sys
from docx import Document

def replace_keyword_in_docx(old_keyword, new_keyword):
    folder_path = os.getcwd()
    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):
            file_path = os.path.join(folder_path, filename)
            doc = Document(file_path)
            for paragraph in doc.paragraphs:
                if old_keyword in paragraph.text:
                    paragraph.text = paragraph.text.replace(old_keyword, new_keyword)
            doc.save(file_path)
            print(f"Processed: {filename}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python change.py <old_keyword> <new_keyword>")
        sys.exit(1)

    old_keyword = sys.argv[1]
    new_keyword = sys.argv[2]

    replace_keyword_in_docx(old_keyword, new_keyword)