import os
import pandas as pd
import numpy as np
from flask import Flask, request, render_template, send_from_directory, flash, redirect, url_for, session
from werkzeug.utils import secure_filename
import datetime
import pickle
import uuid

# --- Configuration ---
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
TEMP_FOLDER = 'temp'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['TEMP_FOLDER'] = TEMP_FOLDER
app.config['SECRET_KEY'] = 'a-very-secret-key' # Change this in production!
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['DOWNLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['TEMP_FOLDER'], exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def add_log(message):
    if 'logs' not in session:
        session['logs'] = []
    session['logs'].append(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}")
    session.modified = True

# --- Business Logic Functions ---
def _create_unique_code(df):
    unique_code_cols = [
        'Sale Document', 'Item (SD)', 'Material', 'Material Description',
        'Plant', 'Storage location', 'Batch'
    ]
    missing_cols = [col for col in unique_code_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Cannot create Unique Item Code. Missing columns: {missing_cols}")
    df['Unique Item Code'] = df[unique_code_cols].astype(str).agg(' | '.join, axis=1)
    return df

def _add_age_columns(df):
    if 'Age in Days' not in df.columns:
        raise ValueError("'Age in Days' column not found for calculating AGE.")
    df['Age in Days'] = pd.to_numeric(df['Age in Days'], errors='coerce').fillna(0)
    df['AGE'] = np.where(df['Age in Days'] >= 150, 'Age >= 150d', 'Age< 150d')
    df['Age <150d'] = df['Age in Days'].apply(lambda x: 'Age < 150d' if x < 150 else '')
    df['Age >=150d'] = df['Age in Days'].apply(lambda x: 'Age >= 150d' if x >= 150 else '')
    return df

def _add_type_responsibility_columns(df):
    if 'Storage location' not in df.columns:
        raise ValueError("'Storage location' column not found for mapping.")
    
    storage_map = {
        '1100': {'Type': 'RM', 'Responsibility': 'PPC'}, 
        '1170': {'Type': 'RM', 'Responsibility': 'Prodn'},
        '1172': {'Type': 'RM', 'Responsibility': 'Prodn'},
        'psa1': {'Type': 'RM', 'Responsibility': 'Prodn'},
        '1111': {'Type': 'SFG', 'Responsibility': 'OID'}, 
        '1112': {'Type': 'SFG', 'Responsibility': 'OID'},
        '1113': {'Type': 'SFG', 'Responsibility': 'OID'}, 
        '1114': {'Type': 'SFG', 'Responsibility': 'OID'},
        '1150': {'Type': 'FG', 'Responsibility': 'Marketing'}, 
        '1173': {'Type': 'FG', 'Responsibility': 'Marketing'},
        '': {'Type': 'FG', 'Responsibility': 'Service'}, 
        '1109': {'Type': 'RAD', 'Responsibility': 'RAD'},
        '1192': {'Type': 'Engg', 'Responsibility': 'Engg'},
        '1193': {'Type': 'Engg', 'Responsibility': 'Engg'}, 
    }

    def get_mapping(loc):
        loc_str = str(loc).strip().lower()
        if loc_str in storage_map:
            return storage_map[loc_str]
        elif 'psa' in loc_str:
            return {'Type': 'RM', 'Responsibility': 'Prodn'}
        else:
            return {'Type': 'FG', 'Responsibility': 'Service'}

    df['Storage location normalized'] = df['Storage location'].astype(str).str.strip().str.lower()
    df['Type'] = df['Storage location normalized'].apply(lambda x: get_mapping(x)['Type'])
    df['Responsibility'] = df['Storage location normalized'].apply(lambda x: get_mapping(x)['Responsibility'])
    df = df.drop(columns=['Storage location normalized'])
    return df

# --- Flask Routes ---
@app.route('/')
def instructions():
    return render_template('instructions.html')

@app.route('/tool')
def tool():
    return render_template('index.html', session=session)

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        flash('No file part', 'warning')
        return redirect(url_for('tool'))
    file = request.files['file']
    if file.filename == '':
        flash('No selected file', 'warning')
        return redirect(url_for('tool'))
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename) # type: ignore
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(upload_path)

        session.clear()
        session['filepath'] = upload_path
        session['filename'] = filename
        add_log(f"File '{filename}' uploaded successfully.")
        session['step'] = 1
        return redirect(url_for('tool'))
    
    flash('Invalid file type. Please upload an Excel file (.xlsx, .xls).', 'warning')
    return redirect(url_for('tool'))

@app.route('/verify')
def verify_file():
    if 'filepath' not in session:
        flash('Please upload a file first.', 'danger')
        return redirect(url_for('tool'))
    try:
        add_log("--- Verifying file... ---")
        df = pd.read_excel(session['filepath'], sheet_name="Total stock", header=1)
        df.columns = [str(col).strip() for col in df.columns]
        
        unique_id = str(uuid.uuid4())
        df_path = os.path.join(app.config['TEMP_FOLDER'], f"{unique_id}.pkl")
        with open(df_path, 'wb') as f:
            pickle.dump(df, f)
        
        session['df_path'] = df_path
        add_log(f"✅ Verification successful. Loaded {len(df)} rows.")
        session['step'] = 2
    except Exception as e:
        add_log(f"❌ Verification failed: {e}")
        flash(f"Verification failed: {e}", 'danger')
    return redirect(url_for('tool'))

@app.route('/find_duplicates')
def find_duplicates():
    if session.get('step', 0) < 2:
        flash('Please verify the file first.', 'danger')
        return redirect(url_for('tool'))
    try:
        with open(session['df_path'], 'rb') as f:
            df = pickle.load(f)
        
        add_log("--- Finding duplicates... ---")
        df_with_code = _create_unique_code(df.copy())
        is_duplicate = df_with_code.duplicated(subset=['Unique Item Code'], keep=False)
        duplicates_df = df_with_code[is_duplicate]

        if duplicates_df.empty:
            add_log("✅ No duplicates were found.")
        else:
            add_log(f"Found {len(duplicates_df)} rows that are part of a duplicate set.")
        
        session['step'] = 3
    except Exception as e:
        add_log(f"❌ Error finding duplicates: {e}")
        flash(f"Error finding duplicates: {e}", 'danger')
    return redirect(url_for('tool'))

@app.route('/process_duplicates')
def process_duplicates():
    if session.get('step', 0) < 3:
        flash('Please find duplicates first.', 'danger')
        return redirect(url_for('tool'))
    try:
        with open(session['df_path'], 'rb') as f:
            df = pickle.load(f)

        add_log("--- Processing duplicates... ---")
        df = _create_unique_code(df)
        is_duplicate = df.duplicated(subset=['Unique Item Code'], keep=False)
        duplicates_df = df[is_duplicate]

        if duplicates_df.empty:
            add_log("✅ No duplicates found to process.")
        else:
            clean_df = df[~is_duplicate].copy()
            sum_cols = ['Unrestricted', 'Value Unrestricted']
            agg_dict = {col: 'sum' for col in sum_cols}
            for col in duplicates_df.columns:
                if col not in sum_cols:
                    agg_dict[col] = 'first'
            aggregated_duplicates = duplicates_df.groupby('Unique Item Code').agg(agg_dict).reset_index(drop=True)
            final_df = pd.concat([clean_df, aggregated_duplicates], ignore_index=True)
            df = final_df.drop(columns=['Unique Item Code'])
            add_log(f"✅ Deduplication process complete. Final row count: {len(df)}")
        
        with open(session['df_path'], 'wb') as f:
            pickle.dump(df, f)

        session['step'] = 4
    except Exception as e:
        add_log(f"❌ Error processing duplicates: {e}")
        flash(f"Error processing duplicates: {e}", 'danger')
    return redirect(url_for('tool'))

@app.route('/create_report')
def create_report():
    if session.get('step', 0) < 4:
        flash('Please process duplicates first.', 'danger')
        return redirect(url_for('tool'))
    try:
        with open(session['df_path'], 'rb') as f:
            final_df = pickle.load(f)

        add_log("--- Creating Final Report... ---")
        final_df = _add_age_columns(final_df)
        final_df = _add_type_responsibility_columns(final_df)

        pivot_df = final_df.pivot_table(
            index=['Type', 'Responsibility'],
            columns='AGE',
            values='Value Unrestricted',
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='Grand Total'
        ).reset_index()

        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        output_filename = f'Total-Stock-Updated-{timestamp}.xlsx'
        output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name='Total Stock', index=False)
            pivot_df.to_excel(writer, sheet_name='Summary', index=False)
        
        add_log(f"✅ Successfully created: {output_filename}")
        session['download_link'] = url_for('download_file', filename=output_filename)
        session['step'] = 5
    except Exception as e:
        add_log(f"❌ Error creating report: {e}")
        flash(f"Error creating report: {e}", 'danger')
    return redirect(url_for('tool'))

@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

@app.route('/reset')
def reset():
    if 'df_path' in session and os.path.exists(session['df_path']):
        os.remove(session['df_path'])
    session.clear()
    flash('Process has been reset.', 'info')
    return redirect(url_for('tool'))

@app.route('/summary_comparator')
def summary_comparator():
    return render_template('summary_comparator.html')

@app.route('/compare_summaries', methods=['POST'])
def compare_summaries():
    if 'file1' not in request.files or 'file2' not in request.files:
        flash('Please select two files.', 'warning')
        return redirect(url_for('summary_comparator'))

    file1 = request.files['file1']
    file2 = request.files['file2']

    if file1.filename == '' or file2.filename == '':
        flash('Please select two files.', 'warning')
        return redirect(url_for('summary_comparator'))

    if file1 and allowed_file(file1.filename) and file2 and allowed_file(file2.filename):
        try:
            def process_summary(df, date):
                df = df.iloc[:-1, :]
                df = df.rename(columns={'Responsibility': 'Responsbilty', 'Grand Total': 'Total',
                                        'Age< 150d': '< 150 Days', 'Age >= 150d': '>= 150 Days'})
                df_long = pd.melt(df, id_vars=['Type', 'Responsbilty'],
                                  value_vars=['Total', '< 150 Days', '>= 150 Days'],
                                  var_name='Metric', value_name=date)
                return df_long

            def format_value(value):
                if pd.isna(value):
                    return ""
                if abs(value) >= 100000:
                    return f'{value / 10000000:.2f} Cr'
                elif abs(value) < 100000:
                    return f'{value / 10000000:.2f}'

            summary_1 = pd.read_excel(file1, sheet_name='Summary')
            summary_2 = pd.read_excel(file2, sheet_name='Summary')

            summary_1.columns = summary_1.columns.str.strip()
            summary_2.columns = summary_2.columns.str.strip()

            summary_1_processed = process_summary(summary_1, 'Date 1')
            summary_2_processed = process_summary(summary_2, 'Date 2')

            comparison_df = pd.merge(summary_1_processed, summary_2_processed, on=['Type', 'Responsbilty', 'Metric'], how='outer')

            comparison_df['Value Change'] = comparison_df['Date 2'] - comparison_df['Date 1']
            comparison_df['Change (%)'] = (comparison_df['Value Change'] / comparison_df['Date 1']).fillna(0)
            
            types = comparison_df['Type'].unique()
            final_df = pd.DataFrame()

            for t in types:
                type_df = comparison_df[comparison_df['Type'] == t]
                total_row = type_df[type_df['Metric'] == 'Total'].groupby('Type').agg({
                    'Date 1': 'sum',
                    'Date 2': 'sum'
                }).reset_index()
                total_row['Responsbilty'] = ''
                total_row['Metric'] = 'Total'
                age_buckets = type_df.groupby(['Type', 'Metric']).agg({
                    'Date 1': 'sum',
                    'Date 2': 'sum'
                }).reset_index()
                age_buckets = age_buckets[age_buckets['Metric'] != 'Total']
                age_buckets['Responsbilty'] = ''

                type_total_df = pd.concat([total_row, age_buckets], ignore_index=True)
                type_total_df['Value Change'] = type_total_df['Date 2'] - type_total_df['Date 1']
                type_total_df['Change (%)'] = (type_total_df['Value Change'] / type_total_df['Date 1']).fillna(0)
                final_df = pd.concat([final_df, type_total_df], ignore_index=True)
                respons_df = type_df[type_df['Metric'] == 'Total'].drop_duplicates(subset=['Type', 'Responsbilty'])
                final_df = pd.concat([final_df, respons_df], ignore_index=True)

            final_df = final_df[['Type', 'Responsbilty', 'Metric', 'Date 1', 'Date 2', 'Value Change', 'Change (%)']]
            final_df = final_df.sort_values(by=['Type', 'Responsbilty']).reset_index(drop=True)

            final_df['Date 1'] = final_df['Date 1'].apply(format_value)
            final_df['Date 2'] = final_df['Date 2'].apply(format_value)
            final_df['Value Change'] = final_df['Value Change'].apply(format_value)
            final_df['Change (%)'] = final_df['Change (%)'].apply(lambda x: f'{x:.2%}')

            timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
            output_filename = f'comparison_summary_{timestamp}.xlsx'
            output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)
            
            final_df.to_excel(output_path, index=False)

            flash('Comparison successful!', 'success')
            return render_template('summary_comparator.html', download_link=url_for('download_file', filename=output_filename))

        except Exception as e:
            flash(f"An error occurred: {e}", 'danger')
            return redirect(url_for('summary_comparator'))

    else:
        flash('Invalid file type. Please upload Excel files only.', 'warning')
        return redirect(url_for('summary_comparator'))

if __name__ == "__main__":
    app.run(debug=True)
