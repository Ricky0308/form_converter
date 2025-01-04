import os
import uuid
import pandas as pd
from markdown import markdown
from weasyprint import HTML
from datetime import datetime
from tkinter import Tk, filedialog, messagebox

def convert_excel_to_pdf(excel_file_path):
    # Step 1: Read the Excel file
    df = pd.read_excel(excel_file_path)

    # Step 2: Create output directory with timestamp
    base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    current_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(current_dir, f"{base_name}_{timestamp}")
    os.makedirs(output_dir, exist_ok=True)

    # Step 3: Process each row
    for idx, row in df.iterrows():
        # Determine the file name
        name = row.get("氏名")
        student_number = row.get("学籍番号")
        if not name and not student_number:
            file_name = f"unknown_{str(uuid.uuid4())[:6]}"
        else:
            file_name = f"{name}_{student_number}"

        sanitized_file_name = str(file_name).replace("/", "-").replace("\\", "-")

        # Create Markdown content
        md_content = """\
"""
        for col_name, cell_value in row.items():
            md_content += f"## {col_name}\n{cell_value if pd.notna(cell_value) else ''}\n\n"

        # Save Markdown file
        md_file_path = os.path.join(output_dir, f"{sanitized_file_name}.md")
        with open(md_file_path, "w", encoding="utf-8") as md_file:
            md_file.write(md_content)

        # Convert Markdown to HTML and then PDF
        html_content = markdown(md_content)
        pdf_file_path = os.path.join(output_dir, f"{sanitized_file_name}.pdf")
        HTML(string=html_content).write_pdf(pdf_file_path)

    messagebox.showinfo("Success", f"PDFs saved in folder: {output_dir}")

def main():
    root = Tk()
    root.withdraw()  # Hide the main Tkinter window

    messagebox.showinfo("Select File", "Please select the Excel file to process.")
    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
    )

    if not excel_file:
        messagebox.showwarning("No File Selected", "No file was selected. Exiting.")
        return

    try:
        convert_excel_to_pdf(excel_file)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    main()
