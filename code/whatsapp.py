"""
Модуль для Whatsapp
"""


import time
from datetime import datetime
import openpyxl
from openpyxl.styles import Font


def clear_name(name):

    name = name.strip()
    name = name.replace('‎', '')
    name = "".join(c for c in name if c.isprintable())
    #name = re.sub(r'\\U.* ', '', name)
    return name


def date_convert(str_date):
    format = '%Y-%m-%d'
    dt = datetime.strptime(str_date, format)
    return dt.timestamp()


def get_time_active(elem):
    if elem[5] >= elem[4] and elem[5] >= elem[3]:
        return("ночная")
    elif elem[4] >= elem[5] and elem[4] >= elem[3]:
        return("вечерняя")
    else:
        return("дневная")


def get_contacts(all):
    contacts = set()
    for elem in all:
        contacts.add(elem[1])
    return contacts


def get_active_contacts(all, t1, t2):
    contacts = dict()

    for elem in all:
        if t1 <= elem[0] <= t2:
            dt = datetime.fromtimestamp(elem[0])
            hour = int(dt.strftime("%H"))
            #print(hour)
            #hour = 1
            if 5 <= hour < 18:
                hour_index = 0
            elif 18 <= hour < 23:
                hour_index = 1
            else:
                hour_index = 2
            if elem[1] in contacts:
                contacts[elem[1]][0] += 1
                if elem[2][0] == '<' and elem[2][len(elem[2]) - 1] == '>':
                    contacts[elem[1]][2] += 1
                else:
                    contacts[elem[1]][1] += len(elem[2])
                contacts[elem[1]][3 + hour_index] += 1
            else:
                #new[0] - all activities, [1] - msgs len, [2] - media number, [3-5] - day time
                new = [1, 0, 0, 0, 0, 0]
                #print(elem[2])
                if elem[2][0] == '<' and elem[2][len(elem[2]) - 1] == '>':
                    new[2] = 1
                else:
                    new[1] = len(elem[2])
                new[3 + hour_index] = 1
                contacts[elem[1]] = new
    return contacts


def read_file(path):
    all = []
    with open(path, 'r', encoding='utf-8', errors='ignore') as file:
        for line in file:
            #line = line.encode("utf-8", 'ignore').decode('utf-8','ignore')
            if len(line) > 20 and line[17:20] == " - ":
                #res = np.arange(3)
                res = [0,'','']
                # Get date
                tmp_list = line.split(" - ")
                d = datetime.strptime(tmp_list[0], "%d.%m.%Y, %H:%M")
                res[0] = time.mktime(d.timetuple())
                # Get name
                tmp_list = tmp_list[1].split(": ")
                # If join group
                if len(tmp_list) == 1:
                    tmp_list = tmp_list[0].split(" вступил")
                    if len(tmp_list) > 1:
                        res[1] = clear_name(tmp_list[0])
                        res[2] = " "
                    else:
                        continue
                # If usual text
                else:
                    res[1] = clear_name(tmp_list[0])
                    res[2] = tmp_list[1].strip()
                if not len(res[2]):
                    res[2] = " "
                if not len(res[1]):
                    res[1] = "<непечатаемые символы>"
                all.append(res)
            # if conitinue of prev string
            else:
                try:
                    all[len(all) - 1][2] += line.strip()
                except:
                    break
    return all


def get_wh_stat(all, file_path, group_name, date1, date2):
    # Excel preparation
    wb = openpyxl.Workbook()
    page = wb.active
    page.column_dimensions['B'].width = 40
    page.column_dimensions['C'].width = 15
    page.column_dimensions['D'].width = 23

    f = open(file_path + '.res', 'w', encoding='utf-8')
    t1 = date_convert(date1)
    t2 = date_convert(date2) + 24 * 3600
    f.write('<h1>%s</h1>Даты с: %s по %s<br>' % (group_name, date1, date2))
    page.append(('', 'Даты с: %s по %s' % (date1, date2)))
    page.append(('', 'Группа: %s' % group_name))
    page.append(('№', 'Имя', 'Показатель', 'Комментарий'))
    page['B1'].font = Font(italic=True)
    page['B2'].font = Font(italic=True)
    page['A3'].font = Font(bold=True)
    page['B3'].font = Font(bold=True)
    page['C3'].font = Font(bold=True)
    page['D3'].font = Font(bold=True)
    a_contacts = get_active_contacts(all, t1, t2)
    sorted_tuples = sorted(a_contacts.items(), key=lambda item: item[1][0], reverse=True)
    sorted_num_dict = {k: v for k, v in sorted_tuples}
    sorted_tuples = sorted(a_contacts.items(), key=lambda item: item[1][1], reverse=True)
    sorted_str_dict = {k: v for k, v in sorted_tuples}
    sorted_tuples = sorted(a_contacts.items(), key=lambda item: item[1][2], reverse=True)
    sorted_media_dict = {k: v for k, v in sorted_tuples}
    action_num = 0
    for key in sorted_num_dict:
        action_num += sorted_num_dict[key][0]

    f.write('Всего активных пользователей в группе: %s<br>' % len(sorted_num_dict))
    f.write('Событий в группе(сообщение, медиафайл, эмоджи): %s' % action_num)
    f.write('%s')
    page.append(('', 'Всего событий в группе:', action_num))
    page.append(('', 'Всего активных пользователей в группе:', len(sorted_num_dict)))

    f.write('<br><br><br><h1>==Активные пользователи==</h1>')
    page.append(('', 'Активные пользователи'))
    line = 6
    page['B%i' % line].font = Font(bold=True)
    i = 1
    for key in sorted_num_dict:
        time_active = get_time_active(sorted_num_dict[key])
        f.write(str(i) + ") " + "%s - %s событий (%s активность)<br>" % (key, str(sorted_num_dict[key][0]), time_active))
        page.append((i, key, sorted_num_dict[key][0], time_active + " активность"))
        line += 1
        i += 1

    f.write('<br><br><h1>==Пользователи с самыми длинными текстами==</h1>')
    page.append(('', 'Пользователи с самыми длинными текстами'))
    line += 1
    page['B%i' % line].font = Font(bold=True)
    i = 1
    for key in sorted_str_dict:
        f.write(str(i) + ") " + "%s - %s символов напечатано<br>" % (key, str(sorted_str_dict[key][1])))
        page.append((i, key, sorted_str_dict[key][1], 'символов напечатано'))
        line += 1
        i += 1
        if i == 11:
            break

    f.write('<br><br><h1>==Пользователи часто использующие медиафайлы==</h1>')
    page.append(('', 'Пользователи часто использующие медиафайлы'))
    line += 1
    page['B%i' % line].font = Font(bold=True)
    i = 1
    for key in sorted_media_dict:
        f.write(str(i) + ") " + "%s - %s медиафайлов отправлено<br>" % (key, str(sorted_media_dict[key][2])))
        page.append((i, key, sorted_media_dict[key][2], 'медиафайлов'))
        line += 1
        i += 1
        if i == 11:
            break
    f.close()
    wb.save(file_path + '.xlsx')
    return file_path + '.res'
