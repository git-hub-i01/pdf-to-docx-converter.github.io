import logging
from flask import Flask, render_template, request, send_from_directory, redirect, url_for, flash
import os
import tempfile
import shutil
import time
import aspose.words as aw
from pdf2docx import Converter

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.secret_key = 'your_secret_key'

# Configure logging
logging.basicConfig(filename='app.log', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file:
            filename = file.filename
            temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(temp_path)
            try:
                if filename.endswith('.pdf'):
                    cv = Converter(temp_path)
                    docx_path = os.path.join(app.config['UPLOAD_FOLDER'], filename[:-4] + '.docx')
                    logging.info('File upload successful: %s', filename)
                    flash('Conversion started...', 'info')
                    start_time = time.time()
                    cv.convert(docx_path)
                    end_time = time.time()
                    conversion_time = end_time - start_time
                    cv.close()
                    flash(f'File converted successfully in {conversion_time:.2f} seconds!', 'success')
                    return redirect(url_for('download_file', name=filename[:-4] + '.docx'))
                elif filename.endswith('.docx'):
                    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename[:-5] + '.pdf')
                    logging.info('File upload successful: %s', filename)
                    flash('Conversion started...', 'info')
                    start_time = time.time()
                    doc = aw.Document(temp_path)
                    doc.save(pdf_path)
                    end_time = time.time()
                    conversion_time = end_time - start_time
                    flash(f'File converted successfully in {conversion_time:.2f} seconds!', 'success')
                    return redirect(url_for('download_file', name=filename[:-5] + '.pdf'))
                elif filename.endswith('.txt'):
                    txt_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    logging.info('File upload successful: %s', filename)
                    flash('Conversion started...', 'info')
                    start_time = time.time()
                    shutil.copy2(temp_path, txt_path)
                    end_time = time.time()
                    conversion_time = end_time - start_time
                    flash(f'File copied successfully in {conversion_time:.2f} seconds!', 'success')
                    return redirect(url_for('download_file', name=filename))
            except Exception as e:
                logging.exception(f"Error converting file: {e}")
                flash(f'Error converting file: {e}', 'error')
                return redirect(url_for('index'))
            finally:
                try:
                    os.remove(temp_path)
                except OSError as e:
                    logging.exception(f"Error deleting temporary file: {e}")
                    print(f"Error deleting temporary file: {e}")
        else:
            flash('Invalid file type. Please upload a PDF file.', 'error')
            return redirect(url_for('index'))
    return render_template('index.html')

@app.route('/docx_to_pdf', methods=['POST'])
def docx_to_pdf():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file and file.filename.endswith('.docx'):
            filename = file.filename
            temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(temp_path)
            try:
                pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename[:-5] + '.pdf')
                logging.info('File upload successful: %s', filename)
                flash('Conversion started...', 'info')
                start_time = time.time()
                doc = aw.Document(temp_path)
                doc.save(pdf_path)
                end_time = time.time()
                conversion_time = end_time - start_time
                flash(f'File converted successfully in {conversion_time:.2f} seconds!', 'success')
                return redirect(url_for('download_file', name=filename[:-5] + '.pdf'))
            except Exception as e:
                logging.exception(f"Error converting file: {e}")
                flash(f'Error converting file: {e}', 'error')
                return redirect(url_for('index'))
            finally:
                try:
                    os.remove(temp_path)
                except OSError as e:
                    logging.exception(f"Error deleting temporary file: {e}")
                    print(f"Error deleting temporary file: {e}")
        else:
            flash('Invalid file type. Please upload a DOCX file.', 'error')
            return redirect(url_for('index'))
    return render_template('index.html')

@app.route('/uploads/<name>')
def download_file(name):
    try:
        return send_from_directory(app.config['UPLOAD_FOLDER'], name, as_attachment=True)
    except FileNotFoundError:
        return "File not found"

if __name__ == '__main__':
    app.run(debug=True)
