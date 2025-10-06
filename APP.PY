import os
import pandas as pd
from flask import Flask, request, render_template, send_file, redirect, url_for, flash, jsonify
from werkzeug.utils import secure_filename

# ---------------------------
# App Config
# ---------------------------
app = Flask(__name__)
app.secret_key = 'supersecretkey'

UPLOAD_FOLDER = 'uploads'
DATABASE_FOLDER = 'database'
OUTPUT_FOLDER = 'outputs'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DATABASE_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ---------------------------
# Temporary storage for TOOL 3
# ---------------------------
groups = {}  # key: filename, value: dict of box_name -> list of row indices

# ---------------------------
# Home Page
# ---------------------------
@app.route('/')
def home():
    return render_template('home.html')

# ======================================================
# TOOL 1: GPON Categorizer
# ======================================================
OUTPUT_FILE = os.path.join(OUTPUT_FOLDER, 'gpon_categorized_output.xlsx')

@app.route('/tool1', methods=['GET', 'POST'])
def tool1():
    if request.method == 'POST':
        files = request.files.getlist('files')
        if len(files) < 2 or len(files) > 3:
            flash('❌ Please upload TWO (Main + Production) or THREE (Main + Production + Lose MAC) files.')
            return redirect(request.url)

        file_paths = []
        for file in files:
            filename = secure_filename(file.filename)
            if not filename:
                continue
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)
            file_paths.append(filepath)

        try:
            result_path, summary = process_files(file_paths)
            return render_template('tool1_result.html', summary=summary, download_link=url_for('download_tool1'))
        except Exception as e:
            flash(f'❌ Error: {str(e)}')
            return redirect(request.url)

    return render_template('tool1.html')

@app.route('/download_tool1')
def download_tool1():
    return send_file(OUTPUT_FILE, as_attachment=True)

def process_files(file_paths):
    def load_file(filepath):
        ext = os.path.splitext(filepath)[-1].lower()
        if ext == '.csv':
            return pd.read_csv(filepath)
        elif ext in ['.xlsx', '.xls']:
            return pd.read_excel(filepath)
        else:
            raise ValueError(f"Unsupported file type: {ext}")

    production_file = next((f for f in file_paths if os.path.basename(f).lower().startswith('prod')), None)
    if not production_file:
        raise ValueError("No production file found (filename should start with 'prod').")

    non_prod_files = [f for f in file_paths if f != production_file]
    if len(non_prod_files) < 1 or len(non_prod_files) > 2:
        raise ValueError("Expected one main file and optionally one lose MAC file.")

    df1 = load_file(non_prod_files[0])
    if len(non_prod_files) == 2:
        df2 = load_file(non_prod_files[1])
        if df1.shape[1] > df2.shape[1]:
            main_file, lose_mac_file = non_prod_files[0], non_prod_files[1]
        else:
            main_file, lose_mac_file = non_prod_files[1], non_prod_files[0]
        lose_mac_df = load_file(lose_mac_file)
    else:
        main_file = non_prod_files[0]
        lose_mac_df = None

    main_df = load_file(main_file)
    production_df = load_file(production_file)

    main_df.columns = [col.lower().strip() for col in main_df.columns]
    production_df.columns = [col.lower().strip() for col in production_df.columns]
    if lose_mac_df is not None:
        lose_mac_df.columns = [col.lower().strip() for col in lose_mac_df.columns]

    gpon_keywords = ['gpon', 'gpon-sn', 'gpon sn']
    mac_keywords = ['mac', 'mac-address', 'mac address']

    main_gpon_col = next((col for col in main_df.columns if any(key in col for key in gpon_keywords)), None)
    main_mac_col = next((col for col in main_df.columns if any(key in col for key in mac_keywords)), None)
    prod_gpon_col = next((col for col in production_df.columns if any(key in col for key in gpon_keywords)), None)

    if not all([main_gpon_col, main_mac_col, prod_gpon_col]):
        raise ValueError("Could not detect GPON or MAC column in main or production file.")

    if lose_mac_df is not None:
        lose_mac_col = next((col for col in lose_mac_df.columns if any(key in col for key in mac_keywords)), None)
        if not lose_mac_col:
            raise ValueError("Could not detect MAC column in Lose MAC file.")

    main_df[main_gpon_col] = main_df[main_gpon_col].astype(str).str.strip().str.upper()
    main_df[main_mac_col] = main_df[main_mac_col].astype(str).str.strip().str.upper()
    production_df[prod_gpon_col] = production_df[prod_gpon_col].astype(str).str.strip().str.upper()

    if lose_mac_df is not None:
        lose_mac_df[lose_mac_col] = lose_mac_df[lose_mac_col].astype(str).str.strip().str.upper()
        gpons_from_lose_macs = main_df[main_df[main_mac_col].isin(lose_mac_df[lose_mac_col])][main_gpon_col].dropna().unique()
    else:
        gpons_from_lose_macs = []

    production_gpons = set(production_df[prod_gpon_col])
    combined_gpons = production_gpons.union(set(gpons_from_lose_macs))

    tested_ok_df = main_df[main_df[main_gpon_col].isin(combined_gpons)]
    rnp_df = main_df[~main_df[main_gpon_col].isin(combined_gpons)]

    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        tested_ok_df.to_excel(writer, sheet_name='Tested OK', index=False)
        rnp_df.to_excel(writer, sheet_name='RNP', index=False)
        if hasattr(gpons_from_lose_macs, 'size') and gpons_from_lose_macs.size > 0:
            pd.DataFrame({main_gpon_col: gpons_from_lose_macs}).to_excel(writer, sheet_name='From Lose MACs', index=False)

    summary = {
        'total_main': len(main_df),
        'tested_ok': len(tested_ok_df),
        'rnp': len(rnp_df),
        'gpons_added_from_mac': len(gpons_from_lose_macs)
    }

    return OUTPUT_FILE, summary

# ======================================================
# TOOL 2: Finder
# ======================================================
@app.route('/tool2', methods=['GET', 'POST'])
def tool2():
    files = [f for f in os.listdir(DATABASE_FOLDER) if f.endswith(('.xlsx', '.xls', '.csv'))]
    data = None
    if request.method == 'POST':
        filename = request.form.get('filename')
        search_value = request.form.get('search_value')

        if filename == '__all__':
            combined_df = pd.DataFrame()
            for f in files:
                path = os.path.join(DATABASE_FOLDER, f)
                if f.endswith('.csv'):
                    df = pd.read_csv(path)
                else:
                    df = pd.read_excel(path)
                combined_df = pd.concat([combined_df, df], ignore_index=True)
            df = combined_df
        else:
            filepath = os.path.join(DATABASE_FOLDER, filename)
            if filename.endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)

        if search_value:
            df = df[df.apply(lambda row: row.astype(str).str.contains(search_value, case=False).any(), axis=1)]

        data = df.to_html(classes="table table-bordered", index=False)

    return render_template('tool2.html', files=files, data=data)

# ======================================================
# TOOL 3: Excel Creator with Box Feature
# ======================================================
@app.route('/tool3', methods=['GET'])
def tool3_page():
    files = [f for f in os.listdir(DATABASE_FOLDER) if f.endswith(('.xlsx','.xls','.csv'))]
    return render_template('tool3.html', files=files, download_link=None)

@app.route('/tool3/check_number', methods=['POST'])
def tool3_check_number():
    try:
        filename = request.form.get('filename')
        number = request.form.get('number').strip()
        box_name = request.form.get('box_name', 'Box 1').strip()
        filepath = os.path.join(DATABASE_FOLDER, filename)

        if filename.endswith('.csv'):
            df = pd.read_csv(filepath, dtype=str)
        else:
            df = pd.read_excel(filepath, dtype=str)

        mask = df.apply(lambda row: row.astype(str).str.strip().eq(number).any(), axis=1)
        if mask.any():
            if filename not in groups:
                groups[filename] = {}
            if box_name not in groups[filename]:
                groups[filename][box_name] = []
            
            new_indices = [i for i in df[mask].index if i not in groups[filename][box_name]]
            groups[filename][box_name].extend(new_indices)
            return jsonify({'status':'found', 'message':f'Number {number} found and added to {box_name}.'})
        else:
            return jsonify({'status':'not_found', 'message':f'Number {number} NOT found in file.'})
    except Exception as e:
        return jsonify({'status':'error', 'message': f'Error checking number: {str(e)}'})

@app.route('/tool3/create_excel', methods=['POST'])
def tool3_create_excel():
    try:
        filename = request.form.get('filename')
        output_name = request.form.get('output_name').strip()
        filepath = os.path.join(DATABASE_FOLDER, filename)

        if filename.endswith('.csv'):
            df = pd.read_csv(filepath, dtype=str)
        else:
            df = pd.read_excel(filepath, dtype=str)

        outpath = os.path.join(OUTPUT_FOLDER, f"{output_name}.xlsx")
        
        with pd.ExcelWriter(outpath, engine='openpyxl') as writer:
            if filename in groups and groups[filename]:
                for box_name, indices in groups[filename].items():
                    if indices:
                        df_filtered = df.loc[indices]
                        # Clean sheet name (Excel has 31 char limit and doesn't allow certain characters)
                        sheet_name = box_name[:31].replace('/', '-').replace('\\', '-')
                        df_filtered.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # Empty file if no data
                pd.DataFrame().to_excel(writer, sheet_name='Empty', index=False)

        download_link = url_for('download_tool3', filename=f"{output_name}.xlsx")

        if filename in groups:
            del groups[filename]

        files = [f for f in os.listdir(DATABASE_FOLDER) if f.endswith(('.xlsx','.xls','.csv'))]
        return render_template('tool3.html', files=files, download_link=download_link)
    except Exception as e:
        return f"Error creating Excel: {str(e)}", 500

@app.route('/download_tool3/<path:filename>')
def download_tool3(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

# ======================================================
# RUN APP
# ======================================================
if __name__ == '__main__':
    app.run(debug=True)