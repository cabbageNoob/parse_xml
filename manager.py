'''
@Descripttion: 
@version: 
@Author: nlpir
@Date: 2019-12-29 20:35:10
@LastEditors  : nlpir
@LastEditTime : 2020-01-03 15:52:58
'''
import os, json
from flask import Flask, flash, render_template, redirect, request, url_for
from werkzeug.utils import secure_filename
from flask import send_from_directory
from utils import util

from xml_process import parse_xml
UPLOAD_FOLDER = 'D:/uploads'
ALLOWED_EXTENSIONS = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif', 'xml'}
# 创建文件目录路径
if not os.path.exists(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['SECRET_KEY'] = '123456'

pwd_path = os.path.abspath(os.path.dirname(__file__))


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


@app.route("/parse_json", methods=["GET", "POST"])
def parse_json():
    file_name = os.path.join(pwd_path, './static/json/')
    if request.method == 'POST':
        time = request.form['time']
        file_name = file_name + time + '.json'
        result = util.readjson(file_name)
        return json.dumps(result)


if __name__ == '__main__':
    app.run('0.0.0.0',port=8001)
