from flask import Flask, request, render_template, send_file
import pytesseract
from PIL import Image
import os
from io import BytesIO
from docx import Document

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'

# Ensure the upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    if file.filename == '':
        return "No selected file"

    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)

        # Perform OCR on the image
        text = pytesseract.image_to_string(Image.open(file_path))

        return render_template('result.html',
                               text=text,
                               filename=file.filename)


@app.route('/download/<filetype>/<filename>', methods=['GET'])
def download_file(filetype, filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    text = pytesseract.image_to_string(Image.open(file_path))

    if filetype == 'txt':
        return send_text_file(text, filename)
    elif filetype == 'docx':
        return send_docx_file(text, filename)
    else:
        return "Invalid file type"


def send_text_file(text, filename):
    buffer = BytesIO()
    buffer.write(text.encode('utf-8'))
    buffer.seek(0)
    return send_file(buffer,
                     as_attachment=True,
                     download_name=f"{os.path.splitext(filename)[0]}.txt",
                     mimetype='text/plain')


def send_docx_file(text, filename):
    document = Document()
    document.add_paragraph(text)
    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{os.path.splitext(filename)[0]}.docx",
        mimetype=
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
