import os

from flask import Flask
from flask import render_template
from flask import request
from flask import send_file
from algo import daily_gl
from algo import file_control
import tarfile

app = Flask(__name__)


@app.route('/', methods=['GET'])
def hello_world():
    return 'Hello World!'


@app.route('/login', methods=['GET', 'POST'])
def render_demo_form():
    if request.method == 'GET':
        return render_template('render_demo_form.html')
    else:
        username = request.form['username']
        password = request.form['password']

        print(username, password)
        return render_template('render_demo.html', username=username, password=password)


@app.route('/excel', methods=['GET'])
def upload_excel_demo_form():
    return render_template('excel_form.html')


@app.route('/excel', methods=['POST'])
def upload_excel_demo_download():

    #empty afe and yongyou
    file_control.delete_direction("afe")
    file_control.delete_direction("yongyou")


    currency_list = request.form['currency_list']
    currency_list = currency_list.split(",")
    date = request.form['date']
    starting_ID = request.form['starting_ID']

    f = request.files['excel']

    print('currency List: %s, filename: %s' % (currency_list, f.filename))

    postfix = f.filename.split('.')[-1]

    file_path = os.getcwd() + '/afe/afe_gl.' + postfix
    f.save(file_path)

    process_gl_data = daily_gl.process_data(file_path)

    for c in currency_list:
        daily_gl.transform_data(process_gl_data, ID=starting_ID, date=date, selector=c)

    tar_file = tarfile.open("%s/yongyou/yongyou_gl.tar.gz" % os.getcwd(), "w:gz")
    tar_file.add("%s/yongyou" % os.getcwd(), arcname="TarName")
    tar_file.close()

    return send_file(os.getcwd() + '/yongyou/' + 'yongyou_gl.tar.gz', as_attachment=True)
    # return render_template('files.html')


@app.route('/download/<name>', methods=['GET'])
def upload_excel_demo_download_2(name):
    return send_file(os.getcwd() + '/tmp/' + name, as_attachment=True)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=9999)
