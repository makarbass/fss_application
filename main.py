from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__)
workbook = load_workbook(filename="z.xlsx")
sheet = workbook.active

#ОТ КОГО 
def fillfio1(i, fio):
    str = [f'R{i}',f'W{i}',f'AB{i}',f'AG{i}',f'AL{i}',f'AQ{i}',f'AV{i}',f'BA{i}',f'BF{i}',f'BK{i}',f'BP{i}',f'BU{i}',f'BZ{i}',f'CE{i}',f'CJ{i}',f'CO{i}',f'CT{i}',f'CY{i}',f'DD{i}',f'DI{i}',f'DN{i}']
    for _ in range (len(fio)):
        sheet[f'{str[_]}'] = fio[_]

# ФИО / РЕГИОН / РАЙОН / ГОРОД / УЛИЦА
def fillfio2(i, fio):
    str = [f'A{i}',f'F{i}',f'K{i}',f'P{i}',f'U{i}',f'Z{i}',f'AE{i}',f'AJ{i}',f'AO{i}',f'AT{i}',f'AY{i}',f'BD{i}',f'BI{i}',f'BN{i}',f'BS{i}',f'BX{i}',f'CC{i}',f'CH{i}',f'CM{i}',f'CR{i}',f'CW{i}',f'DB{i}',f'DG{i}',f'DL{i}']
    for _ in range (len(fio)):
        sheet[f'{str[_]}'] = fio[_]

#ДАТА (РОЖДЕНИЯ/ВЫДАЧИ ПАСПОРТА)
def date(i, data):
    if data:
        data = data.split('-')
        data = data[2] + data[1] + data[0]
    str = [f'A{i}', f'F{i}',f'P{i}', f'U{i}',f'AE{i}', f'AJ{i}',f'AO{i}', f'AT{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

# CНИЛС 
def snils(i, data): # i = 28
    str = [f'C{i}', f'H{i}', f'M{i}', f'W{i}',f'AB{i}',f'AG{i}',f'AQ{i}',f'AV{i}',f'BA{i}',f'BK{i}',f'BP{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

# ИНН
def inn(i, data):
    str = [f'C{i}', f'H{i}', f'M{i}',f'R{i}',f'W{i}',f'AB{i}',f'AG{i}',f'AL{i}',f'AQ{i}',f'AV{i}',f'BA{i}',f'BF{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

# СЕРИЯ
def seria(i, data):
    str = [f'AG{i}',f'AL{i}',f'AQ{i}',f'AV{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

# НОМЕР
def number(i, data):
    str = [f'BN{i}',f'BS{i}',f'BX{i}',f'CC{i}',f'CH{i}',f'CM{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

# КЕМ ВЫДАН / + НАИМЕНОВАНИЕ БАНКА
def threestring(i, data):
    a = 0
    b = 24
    c = i+3
    while i < c:
        fillfio2(i, data[a:b])
        a += 24
        b += 24
        i += 1

# ИНДЕКС
def index(i, data):
    str = [f'A{i}',f'F{i}',f'K{i}',f'P{i}',f'U{i}',f'Z{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

# ДОМ
def house(i, data):
    str = [f'I{i}',f'N{i}',f'S{i}',f'X{i}',f'AC{i}',f'AH{i}',f'AM{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

# КОРПУС
def korp(i, data):
    str = [f'BD{i}',f'BI{i}',f'BN{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

# СТРОЕНИЕ
def stroenie(i, data):
    str = [f'CH{i}',f'CM{i}',f'CR{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

# КВАРТИРА
def kvartira(i, data):
    str = [f'S{i}',f'X{i}',f'AC{i}',f'AH{i}',f'AM{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

# СЧЁТ ПОЛУЧАТЕЛЯ
def schet(i, data):
    str = [f'A{i}',f'F{i}',f'K{i}',f'P{i}',f'U{i}',f'AE{i}',f'AJ{i}',f'AO{i}',f'AT{i}',f'AY{i}',f'BI{i}',f'BN{i}',f'BS{i}',f'BX{i}',f'CC{i}',f'CM{i}',f'CR{i}',f'CW{i}',f'DB{i}',f'DG{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

def bic(i, data):
    str = [f'A{i}',f'F{i}',f'K{i}',f'P{i}',f'U{i}',f'Z{i}',f'AE{i}',f'AJ{i}',f'AO{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

def mir(i, data):
    str = [f'A{i}',f'F{i}',f'K{i}',f'P{i}',f'U{i}',f'Z{i}',f'AE{i}',f'AJ{i}',f'AO{i}',f'AT{i}',f'AY{i}',f'BD{i}',f'BI{i}',f'BN{i}',f'BS{i}',f'BX{i}',f'CC{i}',f'CH{i}',f'CM{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]

def phone(i, data):
    str = [f'AO{i}',f'AY{i}',f'BD{i}',f'BI{i}', f'BS{i}',f'BX{i}',f'CC{i}',f'CM{i}',f'CR{i}', f'DB{i}',f'DG{i}']
    for _ in range (len(data)):
        sheet[f'{str[_]}'] = data[_]


def filldata(data):
    print(data['birth'])
    fillfio1(12, data['surname1'].upper())
    fillfio1(13, data['name1'].upper())
    fillfio1(14, data['second_name1'].upper())
    fillfio2(20, data['surname2'].upper())
    fillfio2(22, data['name2'].upper())
    fillfio2(24, data['second_name2'].upper())
    date(26, data['birth'])
    snils(28, data['snils1']+data['snils2']+data['snils3']+data['snils4'])
    inn(30, data['inn'])
    sheet['N32'] = '✓'
    seria(32, data['seria'])
    number(32, data['number'])
    date(34, data['pdate'])
    threestring(36, data['passportwho'].upper())
    index(71, data['index'])
    fillfio2(73, data['region'].upper())
    fillfio2(75, data['rayon'].upper())
    fillfio2(77, data['city'].upper())
    fillfio2(79, data['street'].upper())
    house(81, data['dom'])
    korp(81, data['corp'])
    stroenie(81, data['stroenie'])
    kvartira(83, data['kvartira'])
    sheet['K103'] = '✓'
    threestring(109, data['bankname'].upper())
    schet(113, data['bs1'])
    bic(115, data['bic'])
    mir(118, data['mir'])
    phone(148, data['phone'])

@app.route("/postData/", methods=['POST'])
def data():
    filldata(dict(request.form))
    file = f"fss_{datetime.now()}.xlsx"
    workbook.save(filename=file)
    return send_file(f'{file}')
    # return request.form

@app.route("/")
def template():
    return render_template("index.html") 

if __name__ == "__main__":
    app.run('10.238.3.216', 5000) 
