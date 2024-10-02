from flask import Flask, request, render_template
from paper_parser import PaperParser
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
parser = PaperParser()

ALLOWED_EXTENSIONS = {'txt', 'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join('/tmp', filename)
            file.save(file_path)
            paper_info = parser.parse_file(file_path)
            os.remove(file_path)  # 删除临时文件
            if paper_info:
                html_content = parser.generate_html(paper_info)
                return render_template('result.html', html_content=html_content)
    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)