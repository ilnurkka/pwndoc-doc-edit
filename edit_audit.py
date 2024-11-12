from flask import Flask, request, send_file
import io
from docx import Document
from preparing import document_preparing

app = Flask(__name__)


@app.route('/edit_docx', methods=['POST'])
def edit_docx():
    file = request.files['file']
    file_content = file.read()
    document = Document(io.BytesIO(file_content))

    document_preparing(document)

    processed_file = io.BytesIO()
    document.save(processed_file)
    processed_file.seek(0)

    return send_file(
        processed_file,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='edited_document.docx'
    )


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
