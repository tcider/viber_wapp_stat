from flask import Flask, request, render_template, url_for
import random
import viber
import whatsapp
from html import unescape
# pip install pysqlite3


TOKEN_SIZE = 10
DB_FOLDER = 'tmp/'
#DB_FOLDER = '/home/a0562265/domains/a0562265.xsph.ru/public_html/tmp/' #FIXME SERVER vs LOCAL
DB_URL = 'http://a0562265.xsph.ru/tmp/'


app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = DB_FOLDER
#app.config['DEBUG'] = True


def generate_token():
    chars = list('abcdefghijklnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890')
    password = ''
    for _ in range(TOKEN_SIZE):
        password += random.choice(chars)
    return password


def get_db_type(token): #FIXME os ext
    file_name_list = token.split(".")
    str_len = len(file_name_list)
    if str_len == 1:
        return 0
    file_ext = file_name_list[str_len - 1]
    if file_ext == 'db':
        return 1
    elif file_ext == 'txt':
        return 2
    else:
        return -1


def get_ftext(type):
    if type == 1:
        return 'Файл с базой Viber загружен'
    elif type == 2:
        return 'Файл с базой WhatsApp загружен'
    elif type == -1:
        return 'Неизвестный тип файла'
    else:
        return 'Файл с Базой не загружен'


def secure_filename(filename):
    filename = filename.replace('"', '').replace("'", '')
    return filename


@app.route("/", methods=['POST', 'GET'])
def index():
    url = url_for('index', _external=True)
    g_text = '<br>'
    date1 = ''
    date2 = ''
    group = ''
    add_text=''
    res_text = 'Для получения статиcтик загрузите файл базы'
    token = generate_token()
    if request.method == 'POST':
        try:
            result = request.form
        except:
            return "Service error"
        token = result['token']
        if 'file_form' in result:
            f = request.files['f']
            file_name = secure_filename(f.filename)
            if len(file_name) and get_db_type(file_name) > 0:
                file_name_list = file_name.split(".")
                if get_db_type(file_name) == 2:
                    group = file_name_list[0]
                file_ext = file_name_list[len(file_name_list) - 1]
                file_name = token + "." + file_ext
                token = file_name

                ##### stream save
                # with open(DB_FOLDER + file_name, "wb") as file:
                #     chunk_size = 4096
                #     while True:
                #         chunk = request.stream.read(chunk_size)
                #         if len(chunk) == 0:
                #             break
                #         file.write(chunk)

                ##### usual save
                f.save(DB_FOLDER + file_name)

        if 'date1' in result and len(result['date1']):
            date1 = result['date1']
        if 'date2' in result and len(result['date2']):
            date2 = result['date2']
        if 'group' in result and len(result['group']):
            group = result['group']


    file_type = get_db_type(token)
    f_text = get_ftext(file_type)
    if file_type == 1:
        add_text = ' повторно'
        groups = viber.get_groups(DB_FOLDER + token)
        if groups:
            g_text = 'Выберите группу:<br><select name="group" size="1">'
            for elem, id in groups:
                selected = ''
                if len(group) and group == f"{id}^{elem}":
                    selected = "selected"
                g_text += f'<option {selected} value="{id}^{elem}">{elem}</option>'
            g_text += '</select><br>'
        else:
            g_text = 'Файл базы Viber не верный.<br>'
        if len(date1) and len(date2) and len(group):
            res_file = viber.get_stat(DB_FOLDER + token, group, date1, date2)
            with open(res_file, "r", encoding='utf-8') as f:
                res_text = f.read()
            link_text = f'<br><a href="{DB_URL}{token}.xlsx">Скачать результат в Excel</a>'
            res_text = res_text % link_text
        else:
            res_text = 'Для получения статиcтик Viber выберите группу и даты'

    elif file_type == 2:
        add_text = ' повторно'
        g_text = f'<input type="hidden" name="group" value="{group}"><br>'
        all = whatsapp.read_file(DB_FOLDER + token)
        res_text = ''
        if len(group):
            res_text = f'<h1>{group}</h1>'
        res_text += 'Для получения статиcтик WhatsApp выберите даты'
        if len(all):
            if len(date1) and len(date2):
                res_file = whatsapp.get_wh_stat(all, DB_FOLDER + token, group, date1, date2)
                with open(res_file, "r", encoding='utf-8') as f:
                    res_text = f.read()
                link_text = f'<br><a href="{DB_URL}{token}.xlsx">Скачать результат в Excel</a>'
                res_text = res_text % link_text
        else:
            g_text += 'Файл базы WhatsApp не верный.<br>'


    return unescape(render_template("index.html", res_text=res_text, f_text=f_text, token=token, url=url, g_text=g_text, \
                                    date1=date1, date2=date2, add_text=add_text))



# @app.route("/upload/<filename>", methods=["POST", "PUT"])
# def upload_process(filename):
#     filename = secure_filename(filename)
#     fileFullPath = os.path.join(application.config['UPLOAD_FOLDER'], filename)
#     with open(fileFullPath, "wb") as f:
#         chunk_size = 4096
#         while True:
#             chunk = flask.request.stream.read(chunk_size)
#             if len(chunk) == 0:
#                 return
#
#             f.write(chunk)
#     return jsonify({'filename': filename})


if __name__ == "__main__":
    app.run()
