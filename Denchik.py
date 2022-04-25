from tkinter import *
from PIL import ImageTk, Image
from tkinter.font import Font
import numpy as np
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

pd.options.mode.chained_assignment = None  # default='warn'
import os
import win32com.client as client
import time


def process_init(init):
    init = init.drop(init.index[2])  # убираем  3 строку с фильтрами (она мешает убирать полностью пустые столбцы)
    init = init.drop(init.columns[init.isna().all()], axis=1)  # удаляем полностью пустые столбцы

    # необходимо поменять шапку таблицы
    init.columns = init.iloc[1, :].values  # заменяем на новую шапку
    init = init.drop(init.index[:2])  # убираем первые две строчки
    indexes = init['var1']
    init = init.drop('var1', 1)  # убираем столбец var1
    init.index = indexes  # заменяем индексы на var1
    return init


def process_db(data):
    # склеиваем код дочки и код услуги в один код
    data['1'] = data['1'].astype(str)
    data['2'] = data['2'].astype(str)
    data['Код ДО-услуга'] = data[['1', '2']].sum(1).astype(int)
    return data


def address_generation(x):  # генерация возможных адресных пространств для лота
    list_of_code = np.array(data['Код ДО-услуга'].drop_duplicates())
    adr = []
    for i in range(len(list_of_code)):
        new = init[init.index == x]['var4'][init[init.index == x][str(list_of_code[i])] == 'ü'].values
        adr.append(new)
    return adr


def create_df_of_adr(data):
    lst = ['AddressSpace0', 'AddressSpace1', 'AddressSpace2', 'AddressSpace3', 'AddressSpace4',
           'AddressSpace5']  # список пространств
    list_of_adr = list(map(address_generation, lst))  # намепливаем функцию на все элементы списка пространств
    list_of_code = np.array(data['Код ДО-услуга'].drop_duplicates())
    dict_of_sp = {}
    for i in range(len(list_of_adr)):
        dict_of_sp['sp' + str(i)] = pd.Series(list_of_adr[i], index=list_of_code)

    df_of_adr = pd.DataFrame(dict_of_sp)  # формируем df адресных пространств по каждому коду
    df_of_adr['Код дочки'] = data[['1', '2']].drop_duplicates()['1'].values
    df_of_adr['Код услуги'] = data[['1', '2']].drop_duplicates()['2'].values

    return df_of_adr


def range_of_date(x):  # геренация промежутков критических значений  для каждого этапа
    for p in list_of_par:
        a = []
        l = list(par.loc[p][x])
        for i in range(len(l)):
            if (i + 1) == par.loc[p][x].shape[0]:
                new = np.arange(l[i], 366)

            else:
                new = np.arange(l[i], l[i + 1])
            a.append(new)
        par[x][p] = a


def create_df_with_params(init):  # создать df критическими значениями
    params = init.loc[list_of_par]
    params = params.drop(params.columns[params.isna().all()], axis=1)  # удаляем полностью пустые столбцы
    params = params.dropna(axis=0, how='any')  # удаляем полностью пустые столбцы
    list_of_levels = init['var3'].loc[list_of_par].drop_duplicates().values
    par = params.drop(['var2', 'var3'], 1)
    par = par.dropna(axis=0, how='all')  # удаляем полностью пустые строки

    return par


def create_df_with_range_of_params(par):  # создать df с промежутками критических значений
    params = init.loc[list_of_par]
    params = params.drop(params.columns[params.isna().all()], axis=1)  # удаляем полностью пустые столбцы
    params = params.dropna(axis=0, how='any')  # удаляем полностью пустые столбцы
    list_of_adr = list(map(range_of_date, par.columns))
    par.drop(par.columns.difference(data['Код ДО-услуга'].values.astype(str)), 1,
             inplace=True)  # оставляем в таблице промежутков значения только для действующих лотов

    par['var3'] = params['var3']  # преобразуем df к удобному для работы виду
    par['var1'] = par.index
    par.index = np.arange(0, len(par))

    return par


def determine_level(x):  # функция определения уровня для каждого лота
    num_of_etap = data['6'][data[data['Код ДО-услуга'] == x[1]].index[0]]  # номер этапа от 1 до 5
    number_of_col = d[num_of_etap]  # номер столбца в data где не nan
    znach = data[number_of_col][data[data['Код ДО-услуга'] == x[1]].index[0]]
    par[str(x[1])][par['var1'] == [d2[num_of_etap]][0]].str.contains(znach,
                                                                     regex=False)  # serias из True и False для lvl
    num_is_true = \
    np.nonzero(np.array(par[str(x[1])][par['var1'] == [d2[num_of_etap]][0]].str.contains(znach, regex=False)))[
        0]  # номер строки где True

    if len(num_is_true) > 0:
        data['LVL'][x[0]] = par['var3'][num_is_true[0]]
    else:
        data['LVL'][x[0]] = 1


def create_df_lv_sp(init, par, data):
    data_with_lvl = data[data['LVL'] != 1]
    list_of_roles = ['Основной', 'Копия', 'Скрытая копия']
    list_of_sp = [[['sp1'], ['sp1', 'sp2'], ['sp3'], ['sp5']], [['sp2'], ['sp3'], ['sp4'], ['sp3', 'sp4']],
                  [['sp0'], ['sp0'], ['sp0'], ['sp0']]]
    list_of_lvl = data_with_lvl['LVL'].drop_duplicates()  # список уровней эскалации
    dict_lvl_sp = {}  # формируем словарь ролей по каждому уровню экалации
    for i in range(len(list_of_roles)):
        dict_lvl_sp[list_of_roles[i]] = pd.Series(list_of_sp[i], index=list_of_lvl.values)
    df_lvl_sp = pd.DataFrame(dict_lvl_sp)  # формируем df ролей по каждому уровню экалации
    return df_lvl_sp


def create_df_with_lvl(init, par):
    data['LVL'] = 0

    list_of_ind_and_cod = np.column_stack([np.arange(0, len(data['Код ДО-услуга'])), data['Код ДО-услуга'].values])
    list_of_lvl = list(map(determine_level, list_of_ind_and_cod))

    data_with_lvl = data[data['LVL'] != 1]  # оставляем те позиции, которые достигли критическихзначений для отправки

    data_with_lvl['Основной'] = 0
    data_with_lvl['Копия'] = 0
    data_with_lvl['Скрытая копия'] = 0
    data_with_lvl['Объединение'] = 0
    data_with_lvl.index = np.arange(0, len(data_with_lvl))

    return data_with_lvl


def fill_adr(x):  # функция заполения ролей адресатов по каждому лоту
    a = []
    df_lvl_sp = create_df_lv_sp(init, par, data)
    list_of_roles = ['Основной', 'Копия', 'Скрытая копия']
    for i in range(len(list_of_roles)):
        lvl = data_with_lvl.loc[x]['LVL']  # определяем уровень
        adr = df_of_adr[df_lvl_sp.loc[lvl][list_of_roles[i]]].loc[
            data_with_lvl['Код ДО-услуга'][x]]  # находим список адресов для данного уровня
        data_with_lvl[list_of_roles[i]][x] = np.concatenate(np.array(adr))
        a.append(list(np.concatenate(np.array(adr))))
    data_with_lvl['Объединение'][x] = str(a)


def create_df_with_full_adr(data_with_lvl):
    list_of_lvl = list(map(fill_adr, data_with_lvl.index))

    data_with_lvl['Объединение'] = data_with_lvl['Объединение'].astype(str)
    data_with_lvl.drop(data_with_lvl[data_with_lvl['Основной'].astype(str) == '[]'].index, inplace=True)

    return data_with_lvl


def send_message(x):
    list_of_columns = ['12', '13', '14', '16', '17', '18', '32', '33', '34', '35',
                       '36']  # выбираем столбцы, которые хотим послать
    new = data_with_lvl_and_adress[data_with_lvl['Объединение'] == x][list_of_columns]  # df для отправки
    new.index = np.arange(0, len(new))
    new = new.dropna(axis=1, how='all')
    new.to_excel("output.xlsx")  # делаем из df файл excel

    path = "C:\\Users\\Nastya\\Desktop\\универ\\Газпромнефть\\output.xlsx"
    a = data_with_lvl_and_adress[data_with_lvl_and_adress['Объединение'] == x]
    a.index = np.arange(0, len(a))
    b = []
    b.append(a['Основной'][0])
    osn = np.array(b).flatten().astype(str)
    cop = np.array(a['Копия'][0]).astype(str)
    hid_cop = np.array(a['Скрытая копия'][0]).astype(str)

    SUBJECT = 'DENCHIK. Информационное сообщение.'  # тема письма
    FILENAME = 'output.xlsx'  # название файла excel
    FILEPATH = 'C:\\Users\\Nastya\\Desktop\\универ\\Газпромнефть\\output.xlsx'  # путь к файлу excel
    MY_EMAIL = 'asdEffff5671@outlook.com'  # логин почты, с которой нужно отправить
    MY_PASSWORD = 'Qwerty12345'  # пароль соответсующей почты
    TO_EMAIL = ';'.join(osn)  # список получателей
    CC = ';'.join(cop)  # копия получателей
    SMTP_SERVER = 'smtp-mail.outlook.com'  # Имя SMTP сервера
    SMTP_PORT = 587  # Порт SMTP

    message = MIMEMultipart()
    message['From'] = MY_EMAIL  # указываем в письмо адрес отправителя
    message['To'] = ''.join(TO_EMAIL)  # указываем в письме адреса получателей
    message['CC'] = ''.join(CC)  # указываем в пиьсме адреса получателей в копии
    message['Subject'] = SUBJECT  # указываем в письме тему

    body_text = "Добрый день,\nАнастасия Эдуардовна.\n\nПо результатам сканирования Вашего плана закупок " \
                "информируем Вас о ближайших событиях, требующих Вашего контроля и внимания\nПросим Вас обеспечить их " \
                "своевременное выполнение или корректрировку плана.\n\nБлагодарим за внимание.\n\nС уважением,\n\nИнформатор " \
                "DENCHIK "  # текстовое наполнение письма

    fp = open(FILENAME, 'rb')  # открытие файла в бинарном режиме
    part = MIMEBase('application', 'vnd.ms-excel')  # заголовок письма
    part.set_payload(fp.read())  # чтение файла
    fp.close  # закрытие файла
    encoders.encode_base64(part)  # шифровка файла под ASCII для отправки по почте
    part.add_header('Content-Disposition', 'attachment', filename=FILENAME)  # указание заголовка для вложения
    message.attach(part)  # прикрипление вложенного файла

    message.attach(MIMEText(body_text, 'plain'))  # прикрипление текста письма

    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.ehlo()  # опредление пользователя на сервере
    server.starttls()  # безопасное соединение
    server.ehlo()
    server.login(MY_EMAIL, MY_PASSWORD)  # верификация почтового аккаунта
    server.sendmail(MY_EMAIL, TO_EMAIL, message.as_string())  # отправка письма
    server.quit()  # выход с сервера
    time.sleep(3)
    #os.remove('output.xlsx')


path_1 = "C:\\Users\\Nastya\\Desktop\\универ\\Газпромнефть\\init2_4.xlsx"
path_2 = "C:\\Users\\Nastya\\Desktop\\универ\\Газпромнефть\\DENCHIKDB.xlsx"

data_1 = pd.read_excel(path_1)  # считываем файл
init = process_init(data_1)

data_2 = pd.read_excel(path_2)  # считываем файл
data = process_db(data_2)

df_of_adr = create_df_of_adr(data)

list_of_par = list(filter(lambda x: x.startswith('Par'), init.index.drop_duplicates()))[
              :-2]  # список этопав на каждом уровне
par = create_df_with_params(init)
par = create_df_with_range_of_params(par)

d = {1: '32', 2: '33', 3: '34', 4: '35', 5: '36'}  # словарь соответствия чисел 1-5 и колонов к в базе
d2 = {}  # словарь соответствия чисел 1-5 и названия этапов
for i in range(len(list_of_par)):
    d2[i + 1] = list_of_par[i]

data_with_lvl = create_df_with_lvl(init, par)

data_with_lvl_and_adress = create_df_with_full_adr(data_with_lvl)

data_with_lvl_and_adress = create_df_with_full_adr(data_with_lvl)

unic_str = list(data_with_lvl_and_adress['Объединение'].drop_duplicates())[
           :2]  # список уникальных строк объединенных адресов


def button_clicked():
    list(map(send_message, unic_str))
    win.quit()



win = Tk()
win.title('Дэнчик')
img = ImageTk.PhotoImage(Image.open("gaz.png"))
imglabel = Label(win, image=img).grid()
button1 = Button(win, text='Начать работу', width=12, height=1, bg='DodgerBlue2', foreground="white", border=0,
                 font=Font(size=20, weight="bold"), command=button_clicked)
button1.grid(row=0, column=0, pady=100, padx=100)
win.mainloop()
