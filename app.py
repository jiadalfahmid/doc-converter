import os
import subprocess
import tempfile
import shutil
import re # Added for file name sanitization
from flask import Flask, request, render_template, send_file, redirect, url_for, flash, after_this_request

# --- Configuration ---
app = Flask(__name__)
app.secret_key = 'super_secret_key_for_flash'  # Required for flash messages

# --- Utility Functions ---

def get_file_extension(filename):
    """Determines the file extension for input."""
    if '.' in filename:
        return filename.rsplit('.', 1)[1].lower()
    return 'txt' # Default to text if no extension

def get_input_format(ext):
    """Maps common extensions to Pandoc input formats. CHANGED TO commonmark for better list support."""
    if ext in ['md', 'markdown']:
        return 'commonmark' # Changed from 'markdown' to 'commonmark'
    elif ext in ['tex', 'latex']:
        return 'latex'
    elif ext in ['html', 'htm']:
        return 'html'
    return 'plain' # Fallback for .txt or unknown

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

    if not input_text and (not uploaded_file or not uploaded_file.filename):
        flash('Please paste text or upload a file before converting.')
        return redirect(url_for('index'))

    # --- Handle Filename Input ---
    output_name_raw = request.form.get('output_filename', 'converted_document').strip()
    
    # Sanitize the filename: remove invalid characters and replace spaces with underscores
    output_name_safe = re.sub(r'[\\/:*?"<>|]', '', output_name_raw)
    output_name_safe = output_name_safe.replace(' ', '_')
    
    # Ensure a default name if sanitization leaves it empty
    if not output_name_safe:
        output_name_safe = 'converted_document'
        
    # Ensure it ends with .docx
    if not output_name_safe.lower().endswith('.docx'):
        output_name_safe += '.docx'
    # -----------------------------------

    # Use mkdtemp() to manually create and manage the temporary directory
    tmpdir = tempfile.mkdtemp()
    
    # Define a cleanup function to run after the file response is sent
    @after_this_request
    def cleanup(response):
        """Schedules the removal of the temporary directory and its contents."""
        try:
            shutil.rmtree(tmpdir)
            print(f"Cleaned up temporary directory: {tmpdir}")
        except Exception as e:
            print(f"Error during temporary directory cleanup {tmpdir}: {e}")
        return response
    
    try:
        input_filename = "input_data"
        input_filepath = os.path.join(tmpdir, input_filename)
        
        # 1. Determine input content and extension
        if uploaded_file and uploaded_file.filename:
            ext = get_file_extension(uploaded_file.filename)
            input_filepath += f".{ext}"
            uploaded_file.save(input_filepath)
        else:
            ext = 'md' 
            input_filepath += f".{ext}"
            with open(input_filepath, 'w', encoding='utf-8') as f:
                f.write(input_text)

        # 2. Configure Pandoc
        input_format = get_input_format(ext)
        output_filename = output_name_safe # Use the safe user-provided name
        output_filepath = os.path.join(tmpdir, output_filename)
        
        pandoc_command = [
            'pandoc',
            input_filepath,
            '-s',
            f'-f{input_format}',
            '-tdocx',
            '-o', output_filepath,
            '--mathml' # Ensures LaTeX equations are rendered correctly within DOCX
        ]

        # 3. Execute Pandoc
        subprocess.run(
            pandoc_command,
            check=True,  # Raise an exception for non-zero exit codes
            capture_output=True,
            text=True
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
    except FileNotFoundError:
        flash('Pandoc is not installed or not found in system PATH. Please check the setup instructions.')
        return redirect(url_for('index'))
    except Exception as e:
        flash(f'An unexpected server error occurred: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    # When deploying, use a production WSGI server like Gunicorn or Waitress.
    app.run(debug=True)
