# -*- coding: utf-8 -*-
"""
Created on Mon Mar 16 10:58:44 2020

@author: пк
"""

from datetime import datetime 

import locale

locale.setlocale(locale.LC_ALL, "") 

import openpyxl as excel

wb_protocol = excel.load_workbook(filename = 'Протокол_комиссии_что_нужно.xlsx')
sheet_protocol = wb_protocol['Протокол']

wb_list = excel.load_workbook(filename = 'Список.xlsx')
sheet_list = wb_list['Лист1']

op = sheet_protocol['A3'].value
kod_op = sheet_protocol['A4'].value
level = sheet_protocol['B6'].value
num_date_protocol = sheet_protocol['A7'].value
discipline = sheet_protocol['A9'].value
student_fio = str(sheet_protocol['B10'].value)
student_group = sheet_protocol['B11'].value
student_numzach = sheet_protocol['B12'].value
course = sheet_protocol['E11'].value
module = sheet_protocol['E12'].value
member1 = sheet_protocol['A14'].value
member2 = sheet_protocol['C14'].value
member3 = sheet_protocol['A15'].value
member4 = sheet_protocol['C15'].value
member5 = sheet_protocol['A16'].value
member6 = sheet_protocol['C16'].value
mark_kom = sheet_protocol['C17'].value
question1 = sheet_protocol['B19'].value
mark1 = sheet_protocol['B22'].value
question2 = sheet_protocol['B23'].value
mark2 = sheet_protocol['B26'].value
question3 = sheet_protocol['B27'].value
mark3 = sheet_protocol['B30'].value
final_comment = sheet_protocol['A31'].value
mark_final = sheet_protocol['B37'].value
mark_final_text = sheet_protocol['D37'].value

for row in sheet_list['A2:T77']:
    # образовательная программа
    sheet_protocol['A3'].value = op + row[0].value 
    # код ОП и направление подготовки
    sheet_protocol['A4'].value = kod_op + row[1].value 
    # уровень образования
    sheet_protocol['B6'].value = row[2].value 
    # номер протокола
    num = row[3].value
    # преобразование даты
    data1 = datetime.strptime(str(row[4].value),'%Y-%m-%d 00:00:00')
    data = data1.strftime('%d %B %Y')
    # ячейка с номером протокола и датой
    sheet_protocol['A7'].value = num_date_protocol[0:num_date_protocol.find('№')+1]+str(num)+num_date_protocol[num_date_protocol.find('\n'):num_date_protocol.find('"')]+'"'+data+'"'
    # Название дисциплины
    sheet_protocol['A9'].value = row[5].value
    # фамилия студента
    student_fio = row[6].value
    sheet_protocol['B10'].value = student_fio
    # группа студента
    sheet_protocol['B11'].value = row[7].value
    # № студенческого
    sheet_protocol['B12'].value = row[8].value
    # курс
    sheet_protocol['E11'].value = course + str(row[9].value)
    # модуль
    sheet_protocol['E12'].value = module + str(row[10].value)
    # Запись членов комиссии, если есть, если нет, то оставляем пустое поле
    if (row[11].value!=None):
        sheet_protocol['A14'].value = member1 + str(row[11].value)
    else:
        sheet_protocol['A14'].value = member1
    if (row[12].value!=None):
        sheet_protocol['C14'].value = member2 + str(row[12].value)
    else:
        sheet_protocol['C14'].value = member2
    if (row[13].value!=None):
        sheet_protocol['A15'].value = member3 + str(row[13].value)
    else:
        sheet_protocol['A15'].value = member3
    if (row[14].value!=None):
        sheet_protocol['C15'].value = member4 + str(row[14].value)
    else:
        sheet_protocol['C15'].value = member4
    if (row[15].value!=None):
        sheet_protocol['A16'].value = member5 + str(row[15].value)
    else:
        sheet_protocol['A16'].value = member5
    if (row[16].value!=None):
        sheet_protocol['C16'].value = member6 + str(row[16].value)
    else:
        sheet_protocol['C16'].value = member6
    # оценка за пересдачу
    sheet_protocol['C17'].value = row[17].value
    
    sheet_protocol['B19'].value = row[20].value #вопрос1
    sheet_protocol['B22'].value = row[21].value #оценка1
    sheet_protocol['B23'].value = row[22].value #вопрос2
    sheet_protocol['B26'].value = row[23].value #оценка2
    sheet_protocol['B27'].value = row[24].value #вопрос3
    sheet_protocol['B30'].value = row[25].value #оценка3
    
    # общий комментарий
    sheet_protocol['A31'].value = row[26].value
    
    # общая оценка
    sheet_protocol['B37'].value = row[18].value
    # общая оценка текстом
    sheet_protocol['D37'].value = row[19].value
    
    # сохранение в отдельный файл
    wb_protocol.save('./Результат/'+student_fio[0:student_fio.find(' ')] + '_Протокол_комиссии'+ str(num) + '_от_' +data1.strftime('%Y%m%d')+'.xlsx')