from flask import Flask, render_template, request, send_file, redirect, url_for, session
import pandas as pd
import os, tempfile
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from PyPDF2 import PdfMerger
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Needed for session
REQUIRED_FIELDS = ["Game_Number", "Gender", "Age", "Date", "Field_Number", "Start_Time",
                   "Home_Team", "Away_Team", "Referee", "Linesman_1", "Linesman_2"]

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/preview-columns', methods=['GET', 'POST'])
def preview_columns():
    if request.method == 'POST':
        file = request.files['file']
        ext = os.path.splitext(file.filename)[1]
        df = pd.read_csv(file) if ext == ".csv" else pd.read_excel(file)

        # Store temp file and dataframe
        temp_dir = tempfile.mkdtemp()
        filename = secure_filename(file.filename)
        filepath = os.path.join(temp_dir, filename)
        file.save(filepath)
        df.to_pickle(os.path.join(temp_dir, "data.pkl"))
        # Save a backup of the original file for reset
        backup_path = os.path.join(temp_dir, f"original{ext}")
        file.stream.seek(0)
        with open(backup_path, 'wb') as f:
            f.write(file.read())
    else:
        temp_dir = request.args.get('temp_dir')
        if not temp_dir or not os.path.exists(os.path.join(temp_dir, "data.pkl")):
            return redirect(url_for('index'))
        df = pd.read_pickle(os.path.join(temp_dir, "data.pkl"))

    df_html = df.to_html(classes="table table-bordered table-sm", index=False)
    return render_template("map_columns.html", columns=df.columns, fields=REQUIRED_FIELDS, temp_dir=temp_dir, table_preview=df_html)

@app.route('/edit-table', methods=['GET', 'POST'])
def edit_table():
    temp_dir = request.args.get('temp_dir') if request.method == 'GET' else request.form.get('temp_dir')
    if not temp_dir or not os.path.exists(os.path.join(temp_dir, "data.pkl")):
        return redirect(url_for('index'))
    df = pd.read_pickle(os.path.join(temp_dir, "data.pkl"))

    # Remove deleted_row and deleted_index logic
    deleted_row = None
    deleted_index = None

    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add_row':
            # Add a new empty row
            new_row = {col: '' for col in df.columns}
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        elif action and action.startswith('delete_row_'):
            # Delete the specified row
            idx = int(action.split('_')[-1])
            df = df.drop(idx).reset_index(drop=True)
        elif action == 'reset_table':
            # Reload the original file and overwrite data.pkl
            import glob
            files = glob.glob(os.path.join(temp_dir, '*'))
            orig_file = [f for f in files if f.endswith('.csv') or f.endswith('.xlsx')]
            if orig_file:
                ext = os.path.splitext(orig_file[0])[1]
                if ext == '.csv':
                    df = pd.read_csv(orig_file[0])
                else:
                    df = pd.read_excel(orig_file[0])
                df.to_pickle(os.path.join(temp_dir, "data.pkl"))
        else:
            # Update DataFrame with submitted data
            new_data = []
            for i in range(len(df)):
                row = []
                for col in df.columns:
                    row.append(request.form.get(f"cell_{i}_{col}"))
                new_data.append(row)
            df = pd.DataFrame(new_data, columns=df.columns)
            df.to_pickle(os.path.join(temp_dir, "data.pkl"))
            return redirect(url_for('preview_columns', temp_dir=temp_dir))
        df.to_pickle(os.path.join(temp_dir, "data.pkl"))
    return render_template("edit_table.html", df=df, temp_dir=temp_dir, deleted_row=deleted_row, deleted_index=deleted_index)

@app.route('/generate-pdf', methods=['POST'])
def generate_pdf():
    temp_dir = request.form['temp_dir']
    df = pd.read_pickle(os.path.join(temp_dir, "data.pkl"))
    mapping = {field: request.form[field] for field in REQUIRED_FIELDS}

    env = Environment(loader=FileSystemLoader("templates"))
    with open("templates/game_card_template.html", "r") as f:
        template = env.from_string(f.read())

    merger = PdfMerger()
    with tempfile.TemporaryDirectory() as tmpdir:
        for i, row in df.iterrows():
            data = {field: str(row[mapping[field]]).strip() for field in REQUIRED_FIELDS}
            html = template.render(**data)
            temp_pdf = os.path.join(tmpdir, f"temp_{i}.pdf")
            HTML(string=html).write_pdf(temp_pdf)
            merger.append(temp_pdf)

        final_pdf = os.path.join(temp_dir, "final_output.pdf")
        merger.write(final_pdf)
        merger.close()

    return send_file(final_pdf, as_attachment=True, download_name="game_cards.pdf")

@app.route('/download-pdf')
def download_pdf():
    temp_dir = request.args.get('temp_dir')
    pdf_path = os.path.join(temp_dir, "final_output.pdf")
    if not os.path.exists(pdf_path):
        return "PDF not found", 404
    return send_file(pdf_path, as_attachment=True, download_name="game_cards.pdf")

if __name__ == '__main__':
    app.run(debug=True)
