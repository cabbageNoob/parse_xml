'''
@Descripttion: 
@version: 
@Author: cjh (492795090@qq.com)
@Date: 2019-12-29 20:35:10
@LastEditors  : cjh (492795090@qq.com)
@LastEditTime : 2019-12-29 20:42:08
'''
import os
from flask import Flask, flash, render_template, redirect, request, url_for
from werkzeug.utils import secure_filename
from flask import send_from_directory

from xml_process import parse_xml
UPLOAD_FOLDER = 'D:/uploads'
ALLOWED_EXTENSIONS = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif', 'xml'}
if not os.path.exists(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SECRET_KEY'] = '123456'


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/parse_xml_api", methods=["GET", "POST"])
def parse_xml_api():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        print(request.files)
        # for file in request.files.getlist('file'):
        #     print('file',file)
        file = request.files['file']
        print('file', file)
        # if user does not select file, browser also
        # submit an empty part without filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            return parse_xml(file_path)
    return render_template("index.html")

if __name__ == '__main__':
    app.run('0.0.0.0',port=8001)
    
