import pandas as pd
import sys
import os
import tempfile
from tkinter import Frame, filedialog, messagebox, Tk, Label, Button, OptionMenu, StringVar
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from PyPDF2 import PdfMerger

REQUIRED_FIELDS = [
    "Gender",
    "Game_Number",
    "Age",
    "Date",
    "Field_Number",
    "Start_Time",
    "Home_Team",
    "Away_Team",
    "Referee",
    "Linesman1",
    "Linesman2",
]

class GameCardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Game Card Generator")
        self.filepath = ""
        self.df = None
        self.dropdown_vars = {}
        self.template_file = os.path.join(os.path.dirname(__file__), "game_card_template.html")

        Button(root, text="Select CSV/XLSX File", command=self.load_file).pack(pady=10)
        self.mapping_frame = Frame(root)
        self.mapping_frame.pack(fill='x', padx=10)
        self.generate_btn = Button(root, text="Generate PDF", command=self.generate_pdf, state='disabled')
        self.generate_btn.pack(pady=10)

    def load_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("CSV and Excel files", "*.csv *.xlsx *.xls")])
        if not self.filepath:
            return
        ext = os.path.splitext(self.filepath)[1].lower()
        try:
            if ext == ".csv":
                self.df = pd.read_csv(self.filepath)
            else:
                self.df = pd.read_excel(self.filepath)
            self.show_dropdowns()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

    def show_dropdowns(self):
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()
        columns = list(self.df.columns)
        self.dropdown_vars = {}
        for idx, field in enumerate(REQUIRED_FIELDS):
            Label(self.mapping_frame, text=field).grid(row=idx, column=0, sticky='w', padx=5, pady=2)
            var = StringVar(self.root)
            var.set(columns[0] if columns else "")
            dropdown = OptionMenu(self.mapping_frame, var, *columns)
            dropdown.grid(row=idx, column=1, sticky='ew', padx=5, pady=2)
            self.dropdown_vars[field] = var
        self.generate_btn['state'] = 'normal'


    def generate_pdf(self):
        mapping = {field: var.get() for field, var in self.dropdown_vars.items()}
        env = Environment(loader=FileSystemLoader('.'))
        with open(self.template_file, "r") as f:
            template = env.from_string(f.read())

        merger = PdfMerger()
        with tempfile.TemporaryDirectory() as tmpdir:
            for i, row in self.df.iterrows():
                try:
                    data = {
                        "Gender": str(row[mapping["Gender"]]).strip(),
                        "Game Number": str(row[mapping["Game_Number"]]).strip(),
                        "Age Group": str(row[mapping["Age"]]).strip(),
                        "Date": str(row[mapping["Date"]]).strip(),
                        "Field Number": str(row[mapping["Field_Number"]]).strip(),
                        "Start Time": str(row[mapping["Start_Time"]]).strip(),
                        "Home Team": str(row[mapping["Home_Team"]]).strip(),
                        "Away Team": str(row[mapping["Away_Team"]]).strip(),
                        "Referee": str(row[mapping["Referee"]]).strip(),
                        "Linesman #1": str(row[mapping["Linesman1"]]).strip(),
                        "Linesman #2": str(row[mapping["Linesman2"]]).strip(),
                    }
                    html = template.render(**data)
                    temp_pdf = os.path.join(tmpdir, f"temp_{i}.pdf")
                    HTML(string=html).write_pdf(temp_pdf)
                    merger.append(temp_pdf)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed on row {i+2}: {e}")
                    return
            merger.write("all_game_cards.pdf")
            merger.close()
        messagebox.showinfo("Success", "PDF created: all_game_cards.pdf")

if __name__ == "__main__":
    root = Tk()
    app = GameCardApp(root)
    root.mainloop()