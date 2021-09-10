"""
Модуль для Viber
"""


import sqlite3
import datetime
import openpyxl
from openpyxl.styles import Font
#pip install openpyxl


def date_convert(str_date):
    format = '%Y-%m-%d'
    dt = datetime.datetime.strptime(str_date, format)
    return dt.timestamp()


def get_name(c, key):
    noname = "Имя не доступно"
    if not key:
        return noname
    c.execute("SELECT Name, ClientName FROM Contact WHERE ContactID = %s" % key)
    name_list = c.fetchall()[0]
    # result = re.sub(r'[^A-zА-я0-9]', '', name)
    # print(result)
    name1 = name_list[0]
    name2 = name_list[1]
    if not name2:
        if not name1:
            name = noname
        else:
            name = name1
    else:
        if not name1:
            name = name2
        else:
            name = name1 + ' (' + name2 + ')'
    name = name.replace('⁨', '').replace('⁩', '')
    return name


def get_stat(file_path, group, date1, date2):
    try:
        conn = sqlite3.connect(file_path)
        c = conn.cursor()
    except:
        return None
    # Excel preparation
    wb = openpyxl.Workbook()
    page = wb.active
    page.column_dimensions['B'].width = 40
    page.column_dimensions['C'].width = 15

    t1 = date_convert(date1) * 1000
    t2 = (date_convert(date2) + 24 * 60 * 60) * 1000
    t1 = str(int(t1))
    t2 = str(int(t2))
    group_list = group.split('^')
    id = group_list[0]
    group = group_list[1]
    c.execute("SELECT count(Events.EventID) FROM Events WHERE TimeStamp > %s AND TimeStamp < %s AND ChatID = %s" % (t1, t2, id))
    res_num = c.fetchall()[0][0]
    f = open(file_path + '.res', 'w', encoding='utf-8')

    f.write('<h1>Группа %s</h1>' % (group))
    f.write('Даты с: %s по %s' % (date1, date2))
    page.append(('', 'Даты с: %s по %s' % (date1, date2)))
    page.append(('', 'Группа: %s' % group))
    page.append(('id', 'Имя', 'Показатель'))
    page['B1'].font = Font(italic=True)
    page['B2'].font = Font(italic=True)
    page['A3'].font = Font(bold=True)
    page['B3'].font = Font(bold=True)
    page['C3'].font = Font(bold=True)
    page.append(('', 'Событий в группе:', res_num))

    c.execute("SELECT ContactID FROM Events WHERE TimeStamp > %s AND TimeStamp < %s AND ChatID = %s" % (t1, t2, id))
    res = c.fetchall()
    dict = {}
    for elem in res:
        contact = elem[0]
        if not contact:
            continue
        if contact in dict:
            dict[contact] += 1
        else:
            dict[contact] = 1
    f.write('<br>Всего активных пользователей в группе: %s</h1>' % len(dict))
    page.append(('', 'Всего активных пользователей в группе:', len(dict)))
    f.write('<br>Событий в группе: %s' % (res_num))
    f.write('%s')
    line = 5
    if len(dict):
        sorted_tuples = sorted(dict.items(), key=lambda item: item[1], reverse=True)
        sorted_dict = {k: v for k, v in sorted_tuples}
        f.write('<br><br><br><h1>==Активные пользователи==</h1>')
        page.append(('', 'Активные пользователи'))
        line += 1
        page['B%i' % line].font = Font(bold=True)
        for key in sorted_dict:
            name = get_name(c, key)
            text2 = 'id - %s, %s, cобытий - %s<br>' % (key, name, str(dict[key]))
            f.write(text2)
            page.append((key, name, dict[key]))
            line += 1
    c.execute("SELECT Events.ContactID FROM Messages, Events WHERE (Messages.Status = 135 OR Messages.Type = 15) AND Messages.EventID = Events.EventID AND Events.TimeStamp > %s AND Events.TimeStamp < %s AND Events.ChatID = %s" % (t1, t2, id))
    # print("SELECT Events.ContactID FROM Messages, Events WHERE (Messages.Status = 135 OR Messages.Type = 15) AND Messages.EventID = Events.EventID AND Events.TimeStamp > %s AND Events.TimeStamp < %s AND Events.ChatID = %s LIMIT 500" % (t1, t2, id))
    res = c.fetchall()
    # print(res)
    res_set = set()
    for elem in res:
        contact = elem[0]
        # if not contact:
        #     continue
        res_set.add(contact)
    f.write('<br><br><h1>==Участники опросов==</h1>')
    page.append(('', 'Участники опросов'))
    line += 1
    page['B%i' % line].font = Font(bold=True)
    for key in res_set:
        name = get_name(c, key)
        f.write('id - %s, %s<br>' % (key, name))
        page.append((key, name))
        line += 1
    if not len(res_set):
        page.append(('', 'нет участников'))
        line += 1
    c.execute("SELECT DISTINCT ContactID FROM Events WHERE ChatID = %s LIMIT 500" % id)
    res = c.fetchall()
    res_set = set()
    for elem in res:
        contact = elem[0]
        if not contact:
            continue
        res_set.add(contact)
    f.write('<br><br><h1>==Все участники сообщества==</h1>')
    page.append(('', 'Все участники сообщества'))
    line += 1
    page['B%i' % line].font = Font(bold=True)
    for key in res_set:
        name = get_name(c, key)
        f.write('id - %s, %s<br>' % (key, name))
        page.append((key, name))
    wb.save(file_path + '.xlsx')
    f.close()
    return file_path + '.res'


def get_groups(file_path):
    try:
        conn = sqlite3.connect(file_path)
        c = conn.cursor()
        c.execute("SELECT Name, ChatID FROM ChatInfo WHERE Name IS NOT NULL")
        res = c.fetchall()
        conn.close()
    except:
        return None
    else:
        return res