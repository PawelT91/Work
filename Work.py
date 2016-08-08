__author__ = 'Тупиков Павел'
__version__ = 1.0

VK = r'\\ilak\database\05-Отчетные\03-01_ГВК МЛ'
GPK_VK = r'\\ilak\database\05-Отчетные\03-03_ГПК МЛ+ПЛ+ПС НКУ'
GPK_VK_O = r'\\ilak\database\05-Отчетные\03-04_ГПК МЛ+ПЛ+ПС НКУ  О'
CKK = r'\\ilak\database\05-Отчетные\03-05_СКК МЛ'
GPK_CI = r'\\ilak\database\05-Отчетные\04-01_ГПК МЛ+ПЛ+ПС МКИ'
GPK_CI_O = r'\\ilak\database\05-Отчетные\04-02_ГПК МЛ+ПЛ+ПС МКИ О'
GPK_MKI = r'\\ilak\database\05-Отчетные\04-03_ГМК МЛ'
NKY = r'\\ilak\database\05-Отчетные\04-05_ГВК МЛ МКИ'
NKY2 = r'\\ilak\database\05-Отчетные\05-01_ГПК МЛ+ПЛ+ПС НКУ'
NKY2_O = r'\\ilak\database\05-Отчетные\05-02_ГПК МЛ+ПЛ+ПС НКУ  О'
CKK_2 = r'\\ilak\database\05-Отчетные\05-03_СКК МЛ'
DATASHEET = r'\\ilak\datasheet\02_Текущие работы'
R = r'\\'

from tkinter import *
import os
from glob import glob
import xlrd
import datetime
import shutil

SELECTED_DOG_VK = ''
SELECTED_NUMBER_VK = ''
SELECTED_DOG_CI = ''
SELECTED_NUMBER_CI = ''
SELECTED_DOG_NKY2 = ''
SELECTED_NUMBER_NKY2 = ''
LIST_VK =[]
LIST_PFK = []
LIST_CI = []
LIST_Lesha = []
LIST_NKY = []
LIST_NKY2 = []

def stop():
    pass

def device():
    os.chdir(r'\\ilak\общая папка\56_Леха Хрисанфов')
    dev = sorted(glob('EQBASE(изм)*'))
    os.startfile(dev[-1])

    
def dog_tec():
    os.chdir(r'\\ilak\личные папки\02_Жаров Алексей Николаевич\Дог_Текущие')
    dev = sorted(glob('Дог_текущие_перечень_*'))
    os.startfile(dev[-1])


    
def datasheet(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    if len(value) != 8:
        i = LIST_VK.index(value)
        print(i)
        while len(LIST_VK[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_VK[gogi],value[2:8]])
        print('---')
        print(x)
        print('---')
        file = (glob(x + '*'))
        print(x[0:6])
        print(x[0:6] ,x[9:13])
        try:
            os.chdir(DATASHEET + R[0] + x[0:6])
            pos = sorted(glob(x[9:13]+'*'))
            print(pos[0])
            os.chdir(DATASHEET + R[0] + x[0:6] + R[0] + pos[0])
            pdf = sorted(glob('*.pdf'))
            os.startfile(pdf[0])
        except:
            pass
    if len(value) == 8:
        os.chdir(r'\\ilak\личные папки\02_Жаров Алексей Николаевич\Дог_Текущие')
        val = sorted(glob(value + '*'))
        print(val)
        
def datasheetPFK(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    if len(value) != 8:
        i = LIST_PFK.index(value)
        print(i)
        while len(LIST_PFK[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_PFK[gogi],value[2:8]])
        print('---')
        print(x)
        print('---')
        file = (glob(x + '*'))
        print(x[0:6])
        print(x[0:6] ,x[9:13])
        try:
            os.chdir(DATASHEET + R[0] + x[0:6])
            pos = sorted(glob(x[9:13]+'*'))
            print(pos[0])
            os.chdir(DATASHEET + R[0] + x[0:6] + R[0] + pos[0])
            pdf = sorted(glob('*.pdf'))
            os.startfile(pdf[0])
        except:
            pass
    if len(value) == 8:
        os.chdir(r'\\ilak\личные папки\02_Жаров Алексей Николаевич\Дог_Текущие')
        val = sorted(glob(value + '*'))
        print(val)
        
def datasheetCI(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    if len(value) != 8:
        i = LIST_CI.index(value)
        print(i)
        while len(LIST_CI[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_CI[gogi],value[2:8]])
        print('---')
        print(x)
        print('---')
        file = (glob(x + '*'))
        print(x[0:6])
        print(x[0:6] ,x[9:13])
        try:
            os.chdir(DATASHEET + R[0] + x[0:6])
            pos = sorted(glob(x[9:13]+'*'))
            print(pos[0])
            os.chdir(DATASHEET + R[0] + x[0:6] + R[0] + pos[0])
            pdf = sorted(glob('*.pdf'))
            os.startfile(pdf[0])
        except:
            pass
    if len(value) == 8:
        os.chdir(r'\\ilak\личные папки\02_Жаров Алексей Николаевич\Дог_Текущие')
        val = sorted(glob(value + '*'))
        print(val)
    

        
def get_pfk():
    os.chdir(GPK_VK)
    file = sorted(glob('*МЛ*058р0.xlsm')) + sorted(glob('*МЛ*058р0_В.xlsm')) 
    number_dog = []
    for f in file:
        if f[0:8] != number_dog and f[0] != '~' and f[0:8] != '149-01-1':
            number_dog = f[0:8]
            listboxPFK.insert(END, number_dog)
            LIST_PFK.append(number_dog)
        if f[0] != '~':
            fil = xlrd.open_workbook(f)
            sheet = fil.sheet_by_index(0)
            val = sheet.row_values(6)[41]
            val2 = sheet.row_values(8)[41]
            val3 = sheet.row_values(7)[41]
            pos = '  ' + f[9:15] + ' - ' + val + ' - ' + val2 + ' - ' + val3
            listboxPFK.insert(END, pos)
            LIST_PFK.append(pos)
            

        
def get_vk():
    os.chdir(VK)
    file = sorted(glob('*_.xlsm'))
    number_dog = []
    for f in file:
        if f[0:8] != number_dog and f[0] != '~' and f[0:8] != '149-01-1':
            number_dog = f[0:8]
            listboxVK.insert(END, number_dog)
            LIST_VK.append(number_dog)
        if f[0] != '~':
            fil = xlrd.open_workbook(f)
            sheet = fil.sheet_by_index(0)
            val = sheet.row_values(6)[41]
            val2 = sheet.row_values(8)[41]
            val3 = sheet.row_values(7)[41]
            pos = '  ' + f[9:15] + ' - ' + val + ' - ' + val2 + ' - ' + val3
            listboxVK.insert(END, pos)
            LIST_VK.append(pos)

def get_ci():
    os.chdir(GPK_CI)
    file = sorted(glob('*МЛ*_058р0_ВС.xlsm')) + sorted(glob('*МЛ*058р0_В.xlsm')) 
    number_dog = []
    for f in file:
        if f[0:8] != number_dog and f[0] != '~' and f[0:8] != '149-01-1':
            number_dog = f[0:8]
            listboxCI.insert(END, number_dog)
            LIST_CI.append(number_dog)
        if f[0] != '~':
            fil = xlrd.open_workbook(f)
            sheet = fil.sheet_by_index(0)
            val = sheet.row_values(6)[41]
            val2 = sheet.row_values(8)[41]
            val3 = sheet.row_values(7)[41]
            pos = '  ' + f[9:15] + ' - ' + val + ' - ' + val2 + ' - ' + val3
            listboxCI.insert(END, pos)
            LIST_CI.append(pos)
            
def get_Lesha():
    os.chdir(GPK_MKI)
    file = sorted(glob('*МЛ*058р0*'))
    number_dog = []
    for f in file:
        if f[0:8] != number_dog and f[0] != '~':
            number_dog = f[0:8]
            listboxLesha.insert(END, number_dog)
            LIST_Lesha.append(number_dog)
        if f[0] != '~':
            fil = xlrd.open_workbook(f)
            sheet = fil.sheet_by_index(0)
            val = sheet.row_values(6)[41]
            val2 = sheet.row_values(8)[41]
            val3 = sheet.row_values(7)[41]
            pos = '  ' + f[9:15] + ' - ' + val + ' - ' + val2 + ' - ' + val3
            listboxLesha.insert(END, pos)
            LIST_Lesha.append(pos)
            
def get_nky():
    os.chdir(NKY)
    file = sorted(glob('*МЛ*058р0*'))
    number_dog = []
    for f in file:
        if f[0:8] != number_dog and f[0] != '~':
            number_dog = f[0:8]
            listboxNKY.insert(END, number_dog)
            LIST_NKY.append(number_dog)
        if f[0] != '~':
            fil = xlrd.open_workbook(f)
            sheet = fil.sheet_by_index(0)
            val = sheet.row_values(6)[41]
            val2 = sheet.row_values(8)[41]
            val3 = sheet.row_values(7)[41]
            pos = '  ' + f[9:15] + ' - ' + val + ' - ' + val2 + ' - ' + val3
            listboxNKY.insert(END, pos)
            LIST_NKY.append(pos)

def get_nky2():
    os.chdir(NKY2)
    file = sorted(glob('*МЛ*058р0*'))
    print (file)
    number_dog = []
    for f in file:
        if f[0:8] != number_dog and f[0] != '~':
            number_dog = f[0:8]
            listboxNKY2.insert(END, number_dog)
            LIST_NKY2.append(number_dog)
        if f[0] != '~':
            fil = xlrd.open_workbook(f)
            sheet = fil.sheet_by_index(0)
            val = sheet.row_values(6)[41]
            val2 = sheet.row_values(8)[41]
            val3 = sheet.row_values(7)[41]
            pos = '  ' + f[9:15] + ' - ' + val + ' - ' + val2 + ' - ' + val3
            listboxNKY2.insert(END, pos)
            LIST_NKY2.append(pos)
    

    
def select(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    print(index,value)
    if len(value) != 8:
        i = LIST_PFK.index(value)
        print(i)
        while len(LIST_PFK[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_PFK[gogi],value[2:8]])
        os.chdir(GPK_VK)
        file = (glob(x + '*'))
        print(file[0])
        global SELECTED_DOG_VK
        global SELECTED_NUMBER_VK
        SELECTED_DOG = LIST_PFK[gogi]
        SELECTED_NUMBER = value[2:8]
        print('Договор ',SELECTED_DOG_VK)
        print('Позиция ',SELECTED_NUMBER_VK)

        
def selectCI(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    print(index,value)
    if len(value) != 8:
        i = LIST_CI.index(value)
        print(i)
        while len(LIST_CI[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_CI[gogi],value[2:8]])
        os.chdir(GPK_CI)
        file = (glob(x + '*'))
        print(file[0])
        global SELECTED_DOG_CI
        global SELECTED_NUMBER_CI
        SELECTED_DOG_CI = LIST_CI[gogi]
        SELECTED_NUMBER_CI = value[2:8]
        print('Договор ',SELECTED_DOG_CI)
        print('Позиция ',SELECTED_NUMBER_CI)

def selectNKY2(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    print(index,value)
    if len(value) != 8:
        i = LIST_NKY2.index(value)
        print(i)
        while len(LIST_NKY2[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_NKY2[gogi],value[2:8]])
        print(x)
        os.chdir(NKY2)
        file = (glob(x + '*'))
        print(file[0])
        global SELECTED_DOG_NKY2
        global SELECTED_NUMBER_NKY2
        SELECTED_DOG_NKY2 = LIST_NKY2[gogi]
        SELECTED_NUMBER_NKY2 = value[2:8]
        print('Договор ',SELECTED_DOG_NKY2)
        print('Позиция ',SELECTED_NUMBER_NKY2)
        
def selectVK(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    rossupT.delete('1.0','12.0')
    finishVK_T.delete('1.0','12.0')
    ingenerVK_T.delete('1.0','12.0')
    print ('You selected item %d: "%s"' % (index, value))
    if len(value) != 8:
        i = LIST_VK.index(value)
        print(i)
        while len(LIST_VK[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_VK[gogi],value[2:8]])
        os.chdir(VK)
        file = (glob(x + '*'))
        print(file[0])
        fil = xlrd.open_workbook(file[0])
        sheet = fil.sheet_by_index(0)
        lenta = sheet.row_values(19)[31]
        data = str(sheet.row_values(16)[43])
        ingerer = sheet.row_values(16)[48]
        print(int(lenta))
        print(ingerer)
        data = int(float(data))
        if data >= 42369:
            year = 2016
            data -= 42369
        if data < 31:
            monch = 1
            data = data 
        elif data < 60:
            monch = 2
            data = data - 31 #29
        elif data < 91:
            monch = 3
            data = data - 60 #31
        elif data < 121:
            monch = 4
            data = data - 91 #30
        elif data < 152:
            monch = 5
            data = data - 121 #31
        elif data < 182:
            monch = 6
            ata = data - 152 #30
        elif data < 213:
            monch = 7
            data = data - 182 #31
        elif data < 244:
            monch = 8
            data = data - 213 #31
        elif data < 274:
            monch = 9
            data = data - 244 #30
        elif data < 305:
            monch = 10
            data = data - 274 #31
        elif data <=335:
            monch = 11
            data = data - 305 #30
        elif data < 366:
            monch = 12
            data = data - 335 #31
        day = data
        date = datetime.date(year,monch,day)
        print(date)
        rossupT.insert(1.0,int(lenta))
        finishVK_T.insert(1.0,date)
        ingenerVK_T.insert(1.0,ingerer)
        stop()

def update():
    listboxPFK.delete(0, END)
    listboxVK.delete(0, END)
    listboxCI.delete(0, END)
    listboxLesha.delete(0, END)
    listboxNKY.delete(0, END)
    listboxNKY2.delete(0, END)
    LIST_VK =[]
    LIST_PFK = []
    LIST_CI = []
    LIST_Lesha = []
    LIST_NKY = []
    LIST_NKY2 = []
    get_vk()
    get_pfk()
    get_ci()
    get_Lesha()
    get_nky()
    get_nky2()
    
    
def select_but():
    LIST_VK =[]
    LIST_PFK = []
    LIST_CI = []
    LIST_Lesha = []
    LIST_NKY = []
    LIST_NKY2 = []
    update()


def open_file(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    if len(value) != 8:
        i = LIST_VK.index(value)
        print(i)
        while len(LIST_VK[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_VK[gogi],value[2:8]])
        print (x)
        os.chdir(VK)
        file = (glob(x + '*'))
        os.startfile(file[0])

def open_filePFK(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    if len(value) != 8:
        i = LIST_PFK.index(value)
        print(i)
        while len(LIST_PFK[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_PFK[gogi],value[2:8]])
        os.chdir(GPK_VK)
        file = (glob(x + '*'))
        for f in file:
            os.startfile(f)
       
def open_fileCI(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    if len(value) != 8:
        i = LIST_CI.index(value)
        print(i)
        while len(LIST_CI[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_CI[gogi],value[2:8]])
        os.chdir(GPK_CI)
        file = (glob(x + '*'))
        for f in file:
            os.startfile(f)
            
def open_fileNKY2(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    if len(value) != 8:
        i = LIST_NKY2.index(value)
        while len(LIST_NKY2[i]) != 8:
            i -= 1
        gogi = i
        x = '-'.join([LIST_NKY2[gogi],value[2:8]])
        print(x)
        os.chdir(NKY2)
        file = sorted(glob(x + '*ПЛ*'))
        if bool(file) is False:
            os.chdir(GPK_CI_O)
            file = sorted(glob(x + '*ПЛ*'))
            print(file)
            if bool(file) is True:
                shutil.copyfile(file[-1], NKY2 + R[0] + file[-1][:-5] + '_КМ_058р0' + file[-1][-5:])
                print(file[-1], 'Скопирован в МКИ')
            os.chdir(NKY2)
        file = (glob(x + '*'))
        for f in file:
            os.startfile(f)
            
def vk_si():
    os.chdir(GPK_VK)
    filePL = sorted(glob(SELECTED_DOG_VK + '-' + SELECTED_NUMBER_VK + '*ПЛ*' + '*'))
    fileML = sorted(glob(SELECTED_DOG_VK + '-' + SELECTED_NUMBER_VK + '*МЛ*' + '*'))
    print(SELECTED_DOG_VK)
    print(SELECTED_NUMBER_VK)
    print(filePL)
    if bool(filePL) is True:
        fi = filePL
        shutil.copyfile(filePL[0], GPK_VK_O + R[0] + filePL[0])
        print(filePL[0], 'Скопирован в отчетные')
        shutil.move(filePL[0], GPK_CI + R[0] + filePL[0][:-5] + 'С' +filePL[0][-5:])
        print(fi, 'Перемещен в МКИ')
    print(fileML)
    if bool(fileML) is True:
        fi = filePL
        shutil.copyfile(fileML[0], GPK_VK_O + R[0] + fileML[0])
        print(fileML[0], 'Скопирован в отчетные')
        shutil.copyfile(fileML[0], CKK + R[0] + fileML[0])
        print(fileML[0], 'Скопирован Наталье Генадьевне')
        shutil.move(fileML[0], GPK_CI + R[0] + fileML[0][:-5] + 'С' + fileML[0][-5:])
        print(fi, 'Перемещен в МКИ')
        update()
        
def vk_NG():
    os.chdir(NKY2)
    filePL = sorted(glob(SELECTED_DOG_NKY2 + '-' + SELECTED_NUMBER_NKY2 + '*ПЛ*' + '*'))
    fileML = sorted(glob(SELECTED_DOG_NKY2 + '-' + SELECTED_NUMBER_NKY2 + '*МЛ*' + '*'))
    print(SELECTED_DOG_NKY2)
    print(SELECTED_NUMBER_NKY2)
    print(filePL)
    if bool(filePL) is True:
        fi = filePL
        shutil.move(filePL[0], NKY2_O + R[0] + filePL[0])
        print(fi, 'Перемещен в отчетные')
    print(fileML)
    if bool(fileML) is True:
        fi = filePL
        shutil.copyfile(fileML[0], NKY2_O + R[0] + fileML[0])
        print(fileML[0], 'Скопирован в отчетные')
        shutil.move(fileML[0], CKK_2 + R[0] + fileML[0])
        print(fileML[0], 'Скопирован Наталье Генадьевне')
        update()

def ci_Lesha():
    os.chdir(GPK_CI)
    filePL = sorted(glob(SELECTED_DOG_CI + '-' + SELECTED_NUMBER_CI + '*ПЛ*' + '*'))
    fileML = sorted(glob(SELECTED_DOG_CI + '-' + SELECTED_NUMBER_CI + '*МЛ*' + '*'))
    print(SELECTED_DOG_CI)
    print(SELECTED_NUMBER_CI)
    print(filePL)
    if bool(filePL) is True:
        fi = filePL
        shutil.move(filePL[0], GPK_CI_O + R[0] + filePL[0])
        print(fi, 'Перемещен в отчетные')
    print(fileML)
    if bool(fileML) is True:
        fi = filePL
        shutil.copyfile(fileML[0], GPK_CI_O + R[0] + fileML[0])
        print(fileML[0], 'Скопирован в отчетные')
        shutil.move(fileML[0], GPK_MKI + R[0] + fileML[0])
        print(fileML[0], 'Скопирован Наталье Генадьевне')
        update()

    
root = Tk()
root.title('РАБОТА')
root.iconbitmap('alfa-komplekt-logo-1458826805.ico')
root.geometry('1700x800')
imageVK = PhotoImage(file = 'ВК.png')
imagePFK = PhotoImage(file = 'ПФК.png')
imageCI = PhotoImage(file = 'CI.png')
imageLesha = PhotoImage(file = 'Lesha.png')
imageNKY = PhotoImage(file = 'NKY.png')
imageNKY2 = PhotoImage(file = 'NKY2.png')
imageFinish = PhotoImage(file = 'Finish.png')
buttonVK = Button(root, image=imageVK, command=select_but)
LabelPFK = Label(root, image=imagePFK)
LabelCI = Label(root, image=imageCI)
LabelLesha = Label(root, image=imageLesha)
LabelNKY = Label(root, image=imageNKY)
LabelNKY2 = Label(root, image=imageNKY2)
LabelFinish = Label(root, image=imageFinish)
listboxVK = Listbox(root, height=15, width=60, selectmode=EXTENDED)
listboxPFK = Listbox(root, height=15, width=60, selectmode=EXTENDED)
listboxCI = Listbox(root, height=15, width=60, selectmode=EXTENDED)
listboxLesha = Listbox(root, height=15, width=60, selectmode=EXTENDED)
listboxNKY = Listbox(root, height=15, width=60, selectmode=EXTENDED)
listboxNKY2 = Listbox(root, height=15, width=60, selectmode=EXTENDED)
rossupT = Text(root,height=0.5,width=12,font='Arial 14',wrap=WORD)
rossupL = Label(root, text = "Количество:")
finishVK_T = Text(root,height=0.5,width=12,font='Arial 14',wrap=WORD)
finishVK_L = Label(root, text = "Окончен ВК:")
ingenerVK_T = Text(root,height=0.5,width=12,font='Arial 14',wrap=WORD)
ingenerVK_L = Label(root, text = "Инженер ВК:")
device = Button(root, text='Приборы',width=10,bg='black',fg='red', command = device)
dog_tec = Button(root, text='Договоры текущие',width=15,bg='blue',fg='yellow', command = dog_tec)
buttonVK_CI = Button(root, text = '>>', command=vk_si)
buttonCI_Lesha = Button(root, text = '>>', command=ci_Lesha)
buttonCI_Finish = Button(root, text = '>>', command=vk_NG)
update()


buttonVK.place(x=80, y=0) # Игорёк))))
listboxVK.place(x=0, y=175)
listboxVK.bind('<<ListboxSelect>>',selectVK)
listboxVK.bind('<Double-Button-1>',open_file)
listboxVK.bind('<Button-3>',datasheet)
listboxPFK.place(x=400, y=175)
listboxPFK.bind('<<ListboxSelect>>',select)
listboxPFK.bind('<Double-Button-1>',open_filePFK)
listboxPFK.bind('<Button-3>',datasheetPFK)
listboxCI.place(x=800, y=175)
listboxCI.bind('<<ListboxSelect>>',selectCI)
listboxCI.bind('<Double-Button-1>',open_fileCI)
listboxCI.bind('<Button-3>',datasheetCI)
listboxNKY2.place(x=400, y=175)
listboxNKY2.bind('<<ListboxSelect>>',selectNKY2)
listboxNKY2.bind('<Double-Button-1>',open_fileNKY2)
listboxLesha.place(x=1200, y=175)
rossupT.place(x=100, y = 420)
rossupL.place(x=0, y = 420)
finishVK_T.place(x=100, y = 450)
finishVK_L.place(x=0, y = 450)
ingenerVK_T.place(x=100, y = 480)
ingenerVK_L.place(x=0, y = 480)
device.place(x=250, y = 420)
dog_tec.place(x=250, y = 450)
LabelPFK.place(x=500, y = 0)
buttonVK_CI.place(x=770, y = 270)
LabelCI.place(x=900, y = 0)
LabelLesha.place(x=1300, y = 0)
buttonCI_Lesha.place(x=1170, y = 270)
LabelNKY.place(x=80, y=520)
listboxNKY.place(x=0, y=695)
LabelNKY2.place(x=500, y=520)
listboxNKY2.place(x=400, y=695)
LabelFinish.place(x=800, y=640)
buttonCI_Finish.place(x=770, y = 770)
root.mainloop()
