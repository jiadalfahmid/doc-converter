import os
import subprocess
import tempfile
import shutil
import re 
import zipfile 
from flask import Flask, request, render_template, send_file, redirect, url_for, flash, after_this_request

# --- Configuration ---
app = Flask(__name__)
app.secret_key = 'super_secret_key_for_flash'

# --- Utility Functions ---

def get_file_extension(filename):
    """Determines the file extension for input."""
    if '.' in filename:
        return filename.rsplit('.', 1)[1].lower()
    return 'txt' 

def get_input_format(ext):
    """Maps common extensions to Pandoc input formats."""
    # Using 'commonmark' for better Markdown list support
    if ext in ['md', 'markdown']:
        return 'commonmark' 
    elif ext in ['tex', 'latex']:
        return 'latex'
    elif ext in ['html', 'htm']:
        return 'html'
    return 'plain'

# --- Routes ---

@app.route('/', methods=['GET'])
def index():
    """Renders the main conversion form."""
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    """Handles the conversion process: takes input, runs Pandoc, returns DOCX."""
    
    input_text = request.form.get('content', '').strip()
    uploaded_file = request.files.get('file')
    # --- NEW: Get selected input format for pasted text
    input_format_selector = request.form.get('input_format_selector', 'md') 

    if not input_text and (not uploaded_file or not uploaded_file.filename):
        flash('Please paste text or upload a file before converting.')
        return redirect(url_for('index'))

    # --- Handle Filename Input ---
    output_name_raw = request.form.get('output_filename', 'converted_document').strip()
    
    # Sanitize the filename
    output_name_safe = re.sub(r'[\\/:*?"<>|]', '', output_name_raw)
    output_name_safe = output_name_safe.replace(' ', '_')
    if not output_name_safe:
        output_name_safe = 'converted_document'
        
    if not output_name_safe.lower().endswith('.docx'):
        output_name_safe += '.docx'
    # -----------------------------

    # Use mkdtemp() for manual temporary directory management
    tmpdir = tempfile.mkdtemp()
    
    # Define a cleanup function to run after the file response is sent
    @after_this_request
    def cleanup(response):
        try:
            shutil.rmtree(tmpdir)
            print(f"Cleaned up temporary directory: {tmpdir}")
        except Exception as e:
            print(f"Error during temporary directory cleanup {tmpdir}: {e}")
        return response
    
    try:
        ext = 'md' # Default extension
        input_filepath = None

        # 1. Determine input content and extension
        if uploaded_file and uploaded_file.filename:
            ext = get_file_extension(uploaded_file.filename)
            
            if ext == 'zip':
                # --- Handle ZIP file containing a full LaTeX project ---
                zip_filepath = os.path.join(tmpdir, "uploaded_project.zip")
                uploaded_file.save(zip_filepath)
                
                project_dir = os.path.join(tmpdir, "project_content")
                os.makedirs(project_dir)
                
                try:
                    with zipfile.ZipFile(zip_filepath, 'r') as zip_ref:
                        zip_ref.extractall(project_dir)
                        
                    input_filename = "main.tex"
                    input_filepath = os.path.join(project_dir, input_filename)
                    
                    if not os.path.exists(input_filepath):
                        flash('ZIP extracted, but "main.tex" not found in root of archive. Please rename your main file to "main.tex" before zipping.')
                        return redirect(url_for('index'))
                    
                    ext = 'tex' # Force extension to LaTeX for Pandoc processing
                    
                except zipfile.BadZipFile:
                    flash("The uploaded file is not a valid ZIP archive.")
                    return redirect(url_for('index'))
            
            else:
                # Handle single file upload (Markdown, HTML, TXT, single .tex)
                input_filename = "input_data"
                input_filepath = os.path.join(tmpdir, input_filename)
                input_filepath += f".{ext}"
                uploaded_file.save(input_filepath)

        else:
            # --- Handle pasted text ---
            ext = input_format_selector # Use the user's selection for the extension
            input_filename = "input_data"
            input_filepath = os.path.join(tmpdir, input_filename)
            input_filepath += f".{ext}"
            with open(input_filepath, 'w', encoding='utf-8') as f:
                f.write(input_text)
        
        if input_filepath is None:
            raise Exception("Input path could not be determined.")

        # 2. Configure Pandoc
        # This will now correctly use 'commonmark' if 'md' is selected, or 'latex' if 'tex' is selected.
        input_format = get_input_format(ext)
        output_filename = output_name_safe
        output_filepath = os.path.join(tmpdir, output_filename)
        
        pandoc_command = [
            'pandoc',
            input_filepath,
            '-s',
            f'-f{input_format}',
            '-tdocx',
            '-o', output_filepath,
            '--mathml' 
        ]
        
        # 3. Execute Pandoc
        cwd = os.path.dirname(input_filepath)
        
        subprocess.run(
            pandoc_command,
            check=True,
            capture_output=True,
            text=True,
            cwd=cwd
        )
            
        # 4. Return the DOCX file
        return send_file(
            output_filepath,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=output_filename
        )

    except subprocess.CalledProcessError as e:
        error_detail = e.stderr.strip().replace('\n', ' ')
        flash(f'Conversion failed (Input format: {ext.upper()}). Error: {error_detail}')
        return redirect(url_for('index'))
    except FileNotFoundError as e:
        if 'main.tex' in str(e):
             flash(str(e))
        else:
            flash('Pandoc is not installed or not found in system PATH. Please check the setup instructions.')
        return redirect(url_for('index'))
    except Exception as e:
        flash(f'An unexpected server error occurred: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
