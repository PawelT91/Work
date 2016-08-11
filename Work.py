__author__ = 'Тупиков Павел'
__version__ = 1.2
# Стандартные библиотеки
from tkinter.ttk import *
from tkinter import *
from glob import glob
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os
import xlrd
import datetime
import shutil
import re
import threading
# Классы
from fileXL import *
from ML import *
# Другие функции
# Папки и катологи
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
# Список сотрудников
LIST_ING_PFK_D = {'Апальков Д.С.': '012','Ломазов Е.Е.': '022','Данилкин П. П.': '041','Михайлюченко И. В.': '045','Шершеньков М. Б.': '049','Тупиков П. А.': '058','Лапидус П.К.': '059','Горчилин В. Ю.': '063','Ларионов В. В.': '071'}
# Выбранный элемент
LEL_DOG = ''
LEL_NOMBER = ''
'''
class Handler(FileSystemEventHandler):
    def on_created(self, event):
        print (event)
        print ('sdfsdfsdfsdfs')

    def on_deleted(self, event):
        print (event)

    def on_moved(self, event):
        print (event)
'''                          
def open_f(dog,nomber):
    file = (glob(dog + '*' + nomber + '*'))
    for f in file:
        os.startfile(f)

def open_file(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    if len(value) != 8:
        nomber = re.findall('\d{4}\.\d{1}', value)[0]
        if (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxVK)):      
            listb = listboxVK
            direct = VK
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxPFK)):
            listb = listboxPFK
            direct = GPK_VK
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxCI)):
            listb = listboxCI
            direct = GPK_CI
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxLesha)):
            listb = listboxLesha
            direct = GPK_MKI
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxNKY)):
            listb = listboxNKY
            direct = NKY
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxNKY2)):
            listb = listboxNKY2
            direct = NKY2
        while len(listb.get(index)) != 8:
            index -= 1
        dog = listb.get(index)
        print(dog,'-',nomber)
        os.chdir(direct)
        open_f(dog,nomber)

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
        nomber = re.findall('\d{4}\.\d{1}', value)[0]
        if (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxVK)):      
            listb = listboxVK
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxPFK)):
            listb = listboxPFK
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxCI)):
            listb = listboxCI
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxLesha)):
            listb = listboxLesha
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxNKY)):
            listb = listboxNKY
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxNKY2)):
            listb = listboxNKY2
        while len(listb.get(index)) != 8:
            index -= 1
        dog = listb.get(index)
        os.chdir(DATASHEET)
        if dog not in glob('*'):
            dog = dog[:-2]
        os.chdir(os.path.join(DATASHEET,dog))
        res = glob(nomber[:-2] + '*')
        if res:
            os.chdir(os.path.join(DATASHEET,dog,res[0]))
            for f in glob('?' + '*.pdf'):
                os.startfile(f)
        else:
            res = glob(nomber[1:-2] + '*')
            if res:
                os.chdir(os.path.join(DATASHEET,dog,res[0]))
                for f in glob('?' + '*.pdf'):
                    os.startfile(f)
             
def get_vk():
    listboxVK.delete(0, END)
    LIST_VK_OBJ = []
    os.chdir(VK)
    file = sorted(glob('*_.xlsm'))
    for f in file:
        obj = ML(f)
        LIST_VK_OBJ.append(obj)
    number_dog = ''
    for f in LIST_VK_OBJ:
        if f.nomber_dog != number_dog:
            number_dog = f.nomber_dog
            listboxVK.insert(END, number_dog)
            listboxVK.itemconfig(listboxVK.size() - 1 , foreground='blue')
        res = '  ' + f.nomber_position + ' - ' + f.device + ' - ' + f.firm + ' - ' + f.device_name
        if f.data_vk_end == '':                                                                           ### !!!! Условие красного цвета
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground='red')
        elif datetime.date.today() - f.data_vk_end >= datetime.timedelta(14):                             ### !!!! Условие красного цвета
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground='red')
        elif datetime.timedelta(14) > datetime.date.today() - f.data_vk_end >= datetime.timedelta(7):     ### !!!! Условие цвета Индиго
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground='indigo')
        else:
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground='green')                                ### !!!! Условие зеленого цвета

def get(lis):
    LIST_OBJ = []
    num = LIST_ING_PFK_D[ComboboxPFK.get()]
    if lis == 'pfk':
        listboxPFK.delete(0, END)
        listb = listboxPFK
        direct = GPK_VK
        os.chdir(direct)
        file = sorted(glob('*МЛ*' + num + 'р' + '?' + '.xlsm')) + sorted(glob('*МЛ*' + num + 'р' + '?' + '_В.xlsm'))
    elif lis == 'ci':
        listboxCI.delete(0, END)
        listb = listboxCI
        direct = GPK_CI
        os.chdir(direct)
        file = sorted(glob('*МЛ*' + num + 'р' + '?' + '_В.xlsm')) + sorted(glob('*МЛ*' + num + 'р' + '?' + '_ВС.xlsm'))
    elif lis == 'Lesha':
        listboxLesha.delete(0, END)
        listb = listboxLesha
        direct = GPK_MKI
        os.chdir(direct)
        file = sorted(glob('*МЛ*' + num + '*' + '.xlsm'))
    elif lis == 'nky':
        listboxNKY.delete(0, END)
        listb = listboxNKY
        direct = NKY
        os.chdir(direct)
        file = sorted(glob('*МЛ*' + num + '*' + '.xlsm'))
    elif lis == 'nky2':
        listboxNKY2.delete(0, END)
        listb = listboxNKY2
        direct = NKY2
        os.chdir(direct)
        file = sorted(glob('*МЛ*' + num + '*' + '.xlsm'))
    if file:
        for f in file:
            if f[0] != '~':
                obj = ML(f)
                LIST_OBJ.append(obj)
    number_dog = ''
    for f in LIST_OBJ:
        if f.nomber_dog != number_dog:
            number_dog = f.nomber_dog
            listb.insert(END, number_dog)
            listb.itemconfig(listb.size() - 1 , foreground='blue')
        res = '  ' + f.nomber_position + ' - ' + f.device + ' - ' + f.firm + ' - ' + f.device_name
        listb.insert(END, res)

def selected(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    if len(value) != 8:
        if (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxPFK)):
            listb = listboxPFK
            os.chdir(GPK_VK)
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxCI)):
            listb = listboxCI
            os.chdir(GPK_CI)
        elif (re.findall('\d{8}', str(event.widget.winfo_screen))[0]) == str(id(listboxNKY2)):
            listb = listboxNKY2
            os.chdir(NKY2)
        while len(listb.get(index)) != 8:
            index -= 1
        nomber = re.findall('\d{4}\.\d{1}', value)[0]
        dog = listb.get(index)
        global LEL_DOG
        global LEL_NOMBER
        LEL_DOG = dog
        LEL_NOMBER = nomber
           
def sel_vk(event):
    LIST_VK = []
    rossupT.delete('1.0','12.0')
    finishVK_T.delete('1.0','12.0')
    engineer_VK_T.delete('1.0','12.0')
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    if len(value) != 8:
        for var in range(listboxVK.size()):
            LIST_VK.append(w.get(var))
        i = LIST_VK.index(value)
        while len(LIST_VK[i]) != 8:
            i -= 1
        dogi = i
        x = '-'.join([LIST_VK[dogi],re.findall('\d{4}\.\d', value)[0]])
        os.chdir(VK)
        file = (glob(x + '*'))
        f = ML(file[0])
        rossupT.insert(1.0,f.quantity)
        finishVK_T.insert(1.0,f.data_vk_end)
        engineer_VK_T.insert(1.0,f.ingerer)

def update_ing(event):
    get('pfk')
    get('ci')
    get('Lesha')
    get('nky')
    get('nky2')
    
def update():
    get_vk()
    update_ing(event = None)
    
def select_but():
    update()
    
def finish():
    if LEL_DOG != '':
        if messagebox.askyesno('Перенос', 'Вы действительно хотите перенести позицию ' + LEL_DOG  + '-' + LEL_NOMBER  + ' на следующий этап?'):
            messagebox.showinfo('Да', 'Хорошо')
            os.chdir(NKY2)
            if len(glob(LEL_DOG + '-' + LEL_NOMBER + '*ПЛ*' + '*')) == 1:
                for pl in sorted(glob( LEL_DOG + '-' + LEL_NOMBER + '*ПЛ*' + '*')):
                    pril = fileXL(pl)
                    pril.move(NKY2_O)
                    print(pril.name_file,'перенесен в отчетные')
            if len(glob( LEL_DOG + '-' + LEL_NOMBER + '*МЛ*' + '*')) == 1:
                for ml in sorted(glob( LEL_DOG + '-' + LEL_NOMBER + '*МЛ*' + '*')):
                    marsh = ML(ml)
                    marsh.copy(NKY2_O)
                    print(marsh.name_file,'перенесен в отчетные')
                    marsh.move(CKK_2)
                    print(marsh.name_file,'перенесена в ССК')
            get('nky2')
            messagebox.showinfo('Перенос', 'Позиция перенесена =)')
        else:
            messagebox.showinfo('Неа', 'Ну и ладно')
        
def finish_vk():
    if LEL_DOG != '':
        if messagebox.askyesno('Перенос', 'Вы действительно хотите перенести позицию ' + LEL_DOG  + '-' + LEL_NOMBER  + ' на следующий этап?'):
            messagebox.showinfo('Да', 'Хорошо')
            os.chdir(GPK_VK)
            if len(glob(LEL_DOG + '-' + LEL_NOMBER + '*ПЛ*' + '*')) == 1:
                for pl in sorted(glob( LEL_DOG + '-' + LEL_NOMBER + '*ПЛ*' + '*')):
                    pril = fileXL(pl)
                    pril.copy(GPK_VK_O)
                    print(pril.name_file,'перенесен в отчетные')
                    pril.move(GPK_CI,'C')
                    print(pril.name_file,'перенесен в CИ')
            if len(glob( LEL_DOG + '-' + LEL_NOMBER + '*МЛ*' + '*')) == 1:
                for ml in sorted(glob( LEL_DOG + '-' + LEL_NOMBER + '*МЛ*' + '*')):
                    marsh = ML(ml)
                    marsh.copy(GPK_VK_O)
                    print(marsh.name_file,'перенесен в отчетные')
                    marsh.copy(GPK_VK_O)
                    print(marsh.name_file,'перенесена в ССК')
                    marsh.move(GPK_CI,'C')
                    print(marsh.name_file,'перенесен в СИ')
            get('pfk')
            messagebox.showinfo('Перенос', 'Позиция перенесена =)')
        else:
            messagebox.showinfo('Неа', 'Ну и ладно')

def finish_ci():
    if LEL_DOG != '':
        if messagebox.askyesno('Перенос', 'Вы действительно хотите перенести позицию ' + LEL_DOG  + '-' + LEL_NOMBER  + ' на следующий этап?'):
            messagebox.showinfo('Да', 'Хорошо')
            os.chdir(GPK_CI)
            if len(glob(LEL_DOG + '-' + LEL_NOMBER + '*ПЛ*' + '*')) == 1:
                for pl in sorted(glob( LEL_DOG + '-' + LEL_NOMBER + '*ПЛ*' + '*')):
                    pril.move(GPK_CI_O)
                    print(pril.name_file,'перенесена в отчетные')
            if len(glob( LEL_DOG + '-' + LEL_NOMBER + '*МЛ*' + '*')) == 1:
                for ml in sorted(glob( LEL_DOG + '-' + LEL_NOMBER + '*МЛ*' + '*')):
                    marsh = ML(ml)
                    marsh.copy(GPK_VK_O)
                    print(marsh.name_file,'перенесен в отчетные')
                    marsh.move(GPK_MKI)
                    print(marsh.name_file,'перенесен к Алексею')
            get('ci')
            messagebox.showinfo('Перенос', 'Позиция перенесена =)')
        else:
            messagebox.showinfo('Неа', 'Ну и ладно')
 
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
ComboboxPFK = Combobox(root, values = list(sorted(LIST_ING_PFK_D.keys())), height=len(list(sorted(LIST_ING_PFK_D.keys()))))
ComboboxPFK.set('Тупиков П. А.')
listboxCI = Listbox(root, height=15, width=60, selectmode=EXTENDED)
listboxLesha = Listbox(root, height=15, width=60, selectmode=EXTENDED)
listboxNKY = Listbox(root, height=15, width=60, selectmode=EXTENDED)
listboxNKY2 = Listbox(root, height=15, width=60, selectmode=EXTENDED)
rossupT = Text(root,height=0.5,width=12,font='Arial 14',wrap=WORD)
rossupL = Label(root, text = "Количество:")
finishVK_T = Text(root,height=0.5,width=12,font='Arial 14',wrap=WORD)
finishVK_L = Label(root, text = "Окончен ВК:")
engineer_VK_T = Text(root,height=0.5,width=12,font='Arial 14',wrap=WORD)
engineer_VK_L = Label(root, text = "Инженер ВК:")
device = Button(root, text='Приборы',width=10,bg='black',fg='red', command = device)
dog_tec = Button(root, text='Договоры текущие',width=15,bg='blue',fg='yellow', command = dog_tec)
buttonVK_CI = Button(root, text = '>>', command=finish_vk)
buttonCI_Lesha = Button(root, text = '>>', command=finish_ci)
buttonCI_Finish = Button(root, text = '>>', command = finish)
update()

listboxVK.place(x=0, y=175)
buttonVK.place(x=80, y=0) # Игорёк))))
listboxVK.place(x=0, y=175)
listboxVK.bind('<<ListboxSelect>>',sel_vk)
listboxVK.bind('<Double-Button-1>',open_file)
listboxVK.bind('<Button-3>',datasheet)
listboxPFK.place(x=400, y=175)
listboxPFK.bind('<<ListboxSelect>>',selected)
ComboboxPFK.bind('<<ComboboxSelected>>', update_ing)
listboxPFK.bind('<Double-Button-1>',open_file)
listboxPFK.bind('<Button-3>',datasheet)
listboxCI.place(x=800, y=175)
listboxCI.bind('<<ListboxSelect>>',selected)
listboxCI.bind('<Double-Button-1>',open_file)
listboxCI.bind('<Button-3>',datasheet)
listboxLesha.place(x=1200, y=175)
listboxLesha.bind('<Double-Button-1>',open_file)
listboxLesha.bind('<Button-3>',datasheet)
listboxNKY.place(x=0, y=695)
listboxNKY.bind('<Double-Button-1>',open_file)
listboxNKY.bind('<Button-3>',datasheet)
listboxNKY2.place(x=400, y=175)
listboxNKY2.bind('<<ListboxSelect>>',selected)
listboxNKY2.bind('<Double-Button-1>',open_file)
listboxNKY2.bind('<Button-3>',datasheet)
rossupT.place(x=100, y = 420)
rossupL.place(x=0, y = 420)
finishVK_T.place(x=100, y = 450)
finishVK_L.place(x=0, y = 450)
engineer_VK_T.place(x=100, y = 480)
engineer_VK_L.place(x=0, y = 480)
device.place(x=250, y = 420)
dog_tec.place(x=250, y = 450)
LabelPFK.place(x=500, y = 0)
ComboboxPFK.place(x=500, y = 450)
buttonVK_CI.place(x=770, y = 270)
LabelCI.place(x=900, y = 0)
LabelLesha.place(x=1300, y = 0)
buttonCI_Lesha.place(x=1170, y = 270)
LabelNKY.place(x=80, y=520)
LabelNKY2.place(x=500, y=520)
listboxNKY2.place(x=400, y=695)
LabelFinish.place(x=800, y=640)
buttonCI_Finish.place(x=770, y = 770)
root.mainloop()
'''
observer = Observer()
observer.schedule(Handler, path=GPK_VK, recursive=True)
observer.start()
try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()
observer.join()
'''
