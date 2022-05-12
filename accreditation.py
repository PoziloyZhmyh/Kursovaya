import os
from openpyxl import load_workbook
from PIL import ImageGrab
import win32com.client
import matplotlib.pyplot as plt
import numpy as np
import tkinter
from tkinter import *
import mysql.connector
import pandas as pd

        #                  Создаем связь с базой данных(Гальчук Георгий)

myconn = mysql.connector.connect(host='cis1.omgtu.ru', user='main_r', passwd='TFADJK2314@$%', db="main_r", port=8306)
cur = myconn.cursor()
trud_df = pd.read_excel('Trudousrt.xlsx', sheet_name='Sheet2')
stud_df = pd.read_excel('Students.xlsx', sheet_name='Sheet 1')
rabotn_df = pd.read_excel('SOTRUDNIKI.xlsx',sheet_name='Sheet1')
prepod_df= pd.read_excel('PREPODY.xlsx',sheet_name='Sheet1')
vypusc_df = pd.read_excel('Vypuskniki.xlsx',sheet_name='Sheet1')

        #                  Считаем баллы вуза по каждому показателю(Ладаев Владислав)
# Считаем средний балл по ЕГЭ
TAB2 = 'select AVG(Ball) from main_r.dbo_studEGE  '
df2 = pd.read_sql(TAB2, con = myconn)
SrEGE=int((df2.iat[0,0])+0.5)

# Считаем средний балл за втупительные испытания
Sum2 = 0
Dlina2 = 0
for i in list(stud_df['FormObuch']):
    for row in list(stud_df['SrVI']):
        if i == "Очно":
            Sum2 = Sum2 + row
            Dlina2 = Dlina2 + 1
        else:
            Sum = Sum2 + 0
    break
SrVI = int((Sum2 / Dlina2)+0.5)

# Считаем процент трудоустроившихся за год
Molodec = 0
NeMolodec=0
for q in list(trud_df['Rabota']):
    if q == "Трудоустроился":
        Molodec = Molodec +1
    else:
        NeMolodec = NeMolodec +1
Rabotaut = int(((Molodec/len(trud_df['Rabota'])) * 100)+0.5)

# Считаем долю обучающихся,успешно завершивших обучение
TAB1 = "select tpID ,appStudStatus from main_r.dbo_StudAppAll where main_r.dbo_StudAppAll.appStudStatus like '%отчислен%' or main_r.dbo_StudAppAll.appStudStatus like '%выпуск%'"
df1 = pd.read_sql(TAB1, con = myconn)
nado=0
for fbb in list(df1['appStudStatus']):
    if fbb == "выпуск":
        nado=nado+1
    else:
        nado=nado
counterFunc = len(df1['appStudStatus'])
DolVypusk=int(((nado/counterFunc)*100)+0.5)

# Считаем долю работников в сфере
TAB4 = "select dolgnost , obraz_full , predstav  from main_r.dbo_pps where main_r.dbo_pps.predstav like ''"
df4 = pd.read_sql(TAB4, con = myconn)
nado4=0
for fff in list(df4['obraz_full']):
    if fff == '':
        nado4=nado4
    else:
        nado4=nado4+1
counterFunc4 = len(df4['obraz_full'])
RabotautVSFERE=int(((nado4/counterFunc4)*100)+0.5)

# Считаем долю сотрудников с ученой степенью
TAB3 = "select wrDOLGNOST ,predstav from main_r.dbo_pps "
df3 = pd.read_sql(TAB3, con = myconn)
nado3=0
for fbf in list(df3['predstav']):
    if fbf == '':
        nado3=nado3
    else:
        nado3=nado3+1
counterFunc3 = len(df3['predstav'])
Uchenye=int(((nado3/counterFunc3)*100)+0.5)

# Наличие внутренней системы оценки качества образования
SysOcenki = True

# Наличие электронной информационно-образовательной среды
ElectrSreda = True

# Доля обучающихся, выполнивших более 70% в диагностической работе
Usp=0
for bet in list(prepod_df['VypAR']):
    if bet == "Не успешно":
        Usp=Usp
    else:
        Usp=Usp+1
Vypolnili = int(((Usp / len(prepod_df['VypAR']))*100)+0.5)

# Выполнение целевого договора
vyp=0
nevyp=0
for a in list(stud_df['VidObuch']):
    for s in list(stud_df['VypUclov']):
        if a == "Целевое":
            if s == "Выполнил":
                vyp = vyp +1
            elif s == "Не выполнил":
                nevyp = nevyp +1
            else:
                vyp = vyp
                nevyp = nevyp
        else:
            vyp = vyp
            nevyp = nevyp
Doly = int(((vyp / (vyp + nevyp))*100)+0.5)

        #                  Расчет набранных баллов(Ладаев Владислав)

# Расчет суммы баллов 1-ой группы показателей
if SrEGE >= 66:
    Sum1_1=  10
elif SrEGE>=60 and SrEGE<=65:
    Sum1_1 = 5
else:
    Sum1_1= 0

if SrVI >= 66:
    Sum1_2= 10
elif SrVI>=60 and SrVI<=65:
    Sum1_2 = 5
else:
    Sum1_2= 0

if ElectrSreda == True:
    Sum1_3 = 10
else:
    Sum1_3 = 0

if Uchenye >= 60:
    Sum1_4= 20
elif Uchenye>=50 and Uchenye<=59:
    Sum1_4 = 5
else:
    Sum1_4 = 0

if RabotautVSFERE >= 70:
    Sum1_5 = 20
else:
    Sum1_5 = 0

if Vypolnili >= 65:
    Sum1_6 = 75
elif Vypolnili>=55 and Vypolnili<=64:
    Sum1_6 = 40
else:
    Sum1_6 = 0

if SysOcenki == True:
    Sum1_7 = 10
else:
    Sum1_7 = 0

ObshBall1 = Sum1_1 + Sum1_2 + Sum1_3 + Sum1_4 + Sum1_5 + Sum1_6 + Sum1_7
f1='a'
if ObshBall1>=70:
    f1='Вуз получил государственную аккредитацию профессиональной деятельности,тк набрал: '
else:
    f1='Вуз не получил государственную аккредитацию профессиональной деятельности,тк набрал: '

# Расчет суммы баллов 2-ой группы показателей
if SrEGE >= 66:
    Sum2_1= 10
elif SrEGE>=60 and SrEGE<=65:
    Sum2_1 = 5
else:
    Sum2_1 = 0

if SrVI >= 66:
    Sum2_2 = 10
elif SrVI>=60 and SrVI<=65:
    Sum2_2 = 5
else:
    Sum2_2 = 0

if ElectrSreda == True:
    Sum2_3 = 10
else:
    Sum2_3 = 0

if DolVypusk >= 70:
    Sum2_4 = 10
elif DolVypusk>=50 and SrVI<=69:
    Sum2_4 = 5
else:
    Sum2_4 = 0

if Doly >= 50:
    Sum2_5 = 10
elif Doly>=30 and Doly<=49:
    Sum2_5 = 5
else:
    Sum2_5 = 0

if Uchenye >= 60:
    Sum2_6 = 20
elif Uchenye>=50 and Uchenye<=59:
    Sum2_6 = 5
else:
    Sum2_6 = 0

if RabotautVSFERE >= 70:
    Sum2_7 = 20
else:
    Sum2_7 = 0

if SysOcenki == True:
    Sum2_8 = 10
else:
    Sum2_8 = 0

if Rabotaut >= 75:
    Sum2_9 = 20
elif Doly>=50 and Doly<=74:
    Sum2_9 = 10
else:
    Sum2_9 = 0

ObshBall2 = Sum2_1 + Sum2_2 + Sum2_3 + Sum2_4 + Sum2_5 + Sum2_6 + Sum2_7 + Sum2_8 + Sum2_9
f2='a'
if ObshBall2>=70:
    f2='Вуз получил аккредитацию по результатам мониторинга,тк набрал: '
else:
    f2='Вуз не получил аккредитацию по результатам мониторинга,тк набрал: '

# Расчет суммы баллов 3-ей группы показателей
if Vypolnili >= 65:
    Sum3_1 = 75
elif Vypolnili>=55 and Vypolnili<=64:
    Sum3_1 = 40
else:
    Sum3_1 = 0

if SysOcenki == True:
    Sum3_2 = 20
else:
    Sum3_2 = 0

ObshBall3 = Sum3_1 + Sum3_2
f3='a'
if ObshBall3>=60:
    f3='Вуз получил аккредитацию по результатам федерального государственного контроля,тк набрал: '
else:
    f3='Вуз не получил аккредитацию по результатам федерального государственного контроля,тк набрал: '

        #                  Создаем графический интерфейс пользователя(Гальчук Георгий)

# Строим диаграмму для наглядности представления информации
labels = ['1 группа\nпоказателей', '2 группа\nпоказателей', '3 группа\nпоказателей']
NABRAL = [ObshBall1,ObshBall2,ObshBall3]
NADO = [90,70,60]

x = np.arange(len(labels))  #
width = 0.35

fig, ax = plt.subplots()
rects1 = ax.bar(x - width/2, NABRAL, width, label='Набранные баллы')
rects2 = ax.bar(x + width/2, NADO, width, label='Проходной минимум')

ax.set_ylabel('Баллы')
ax.set_title('Аккредитация вуза')
ax.set_xticks(x, labels)
ax.legend()

ax.bar_label(rects1, padding=3)
ax.bar_label(rects2, padding=3)

fig.tight_layout()
fig.savefig('graf.png')

# Создаем рабочее окно
root = tkinter.Tk()
root.title("Расчет аккредитационных показателей вуза")
# создаем рабочую область
frame = tkinter.Frame(root)
frame.grid()

#Добавим метку
label = tkinter.Label(frame, text="В результате проведенных расчетов получили следующие результаты:",bg='white').grid(row=1,column=1)

#Добавим изображение
canvas = tkinter.Canvas(root, height=480, width=640,bg='white')
img = tkinter.PhotoImage(file = 'graf.png')
image = canvas.create_image(0, 0, anchor='nw',image=img)
canvas.grid(row=6,column=0)
#Добавим пояснения
lbl=Label(text=f1+str(ObshBall1)+' баллов.',bg='white').grid(row=2,column=0)
lbl2=Label(text=f2+str(ObshBall2)+' баллов.',bg='white').grid(row=3,column=0)
lbl3=Label(text=f3+str(ObshBall3)+' баллов.',bg='white').grid(row=4,column=0)

root["bg"] = "white"
root.geometry('700x600')
root.resizable(width=False, height=False)

#Создадим таблицу с результатами вычисления
fn = 'sa.xlsx'
wb0 = load_workbook(fn)
ws0 = wb0.active
ws0['C2']= Sum1_1
ws0['C3']= Sum1_2
ws0['C4']= Sum1_3
ws0['C5']= Sum1_4
ws0['C6']= Sum1_5
ws0['C7']= Sum1_6
ws0['C8']= Sum1_7
ws0['C9']= ObshBall1
ws0['C11']= Sum2_1
ws0['C12']= Sum2_2
ws0['C13']= Sum2_3
ws0['C14']= Sum2_4
ws0['C15']= Sum2_5
ws0['C16']= Sum2_6
ws0['C17']= Sum2_7
ws0['C18']= Sum2_8
ws0['C19']= Sum2_9
ws0['C20']= ObshBall2
ws0['C22']= Sum3_1
ws0['C23']= Sum3_2
ws0['C24']= ObshBall3
wb0.save(fn)
wb0.close()

#Добавим таблицу с пояснениями во 2 рабочее окно
root_path = os.path.dirname(os.path.abspath(__file__))
def main():
    def get_path(*path):
        return os.path.join(root_path, *path)

    xlsx_path = get_path('sa.xlsx')
    client = win32com.client.Dispatch("Excel.Application")

    wb = client.Workbooks.Open(xlsx_path)
    ws = wb.Worksheets("Sheet1")

    def fpart():
        ws.Range("A1:C9").CopyPicture(Format = 2)
        img0 = ImageGrab.grabclipboard()
        img0.save(get_path('fg1.png'))
    def spart():
        ws.Range("A10:C20").CopyPicture(Format = 2)
        img0 = ImageGrab.grabclipboard()
        img0.save(get_path('fg2.png'))
    def tpart():
        ws.Range("A21:C24").CopyPicture(Format = 2)
        img0 = ImageGrab.grabclipboard()
        img0.save(get_path('fg3.png'))

    fpart()
    spart()
    tpart()

    wb.Close()
    client.Quit()
main()

# Свяжем 1 и 2 рабочие окна
def createNewWindow():
    root1 = Toplevel()
    root1.geometry('1150x850')

    canvas1 = tkinter.Canvas(root1, height=263, width=1102, bg='white')
    img1 = tkinter.PhotoImage(file='fg1.png')
    image1 = canvas1.create_image(0, 0, anchor='nw', image=img1)
    canvas1.grid(row=1, column=0)
    lbl1 = Label(root1,text='Для целей государственной аккредитации образовательной деятельности (минимальное значение 90 баллов).',bg='white',font="Arial 12").grid(row=0, column=0)

    canvas2 = tkinter.Canvas(root1, height=311, width=1102, bg='white')
    img2 = tkinter.PhotoImage(file='fg2.png')
    image2 = canvas2.create_image(0, 0, anchor='nw', image=img2)
    canvas2.grid(row=3, column=0)
    lbl2 = Label(root1,text='Для целей осуществления аккредитационного мониторинга (минимальное значение 70 баллов).',bg='white',font="Arial 12").grid(row=2, column=0)

    canvas3 = tkinter.Canvas(root1, height=238, width=1102, bg='white')
    img3 = tkinter.PhotoImage(file='fg3.png')
    image3 = canvas3.create_image(0, 0, anchor='nw', image=img3)
    canvas3.grid(row=5, column=0)
    lbl3 = Label(root1,text='Для целей осуществления федерального государственного контроля (надзора) в сфере образования (минимальное значение 60 баллов).',bg='white',font="Arial 12").grid(row=4, column=0)

    root1.mainloop()

# вставляем кнопку
but = tkinter.Button( text="Набранные баллы по каждому аккредитационному показателю",command=createNewWindow).grid(row=10, column=0)

# Получаем готовый интерфейс пользователя
root.mainloop()


