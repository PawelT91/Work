__author__ = 'Тупиков Павел'
__version__ = 1.4
# Стандартные библиотеки
from tkinter.ttk import *
from tkinter import *
from glob import glob
#from watchdog.observers import Observer
#from watchdog.events import FileSystemEventHandler
import os
import xlrd
import datetime
import shutil
import re
import threading
import time
import threading
# Классы
class fileXL:
    def __init__(self,name_file):
        self.name_file = name_file
        self.nomber_dog = re.findall('.{3}\-.{2}\-.{1}',name_file)[0]
        self.nomber_position = re.findall('.{4}\.\d+', name_file)[0]

    def copy(self, direct, char = ''):
        shutil.copyfile(self.name_file, os.path.join(direct,self.name_file[:-5] + char + self.name_file[-5:]))

    def move(self, direct, char = ''):
        shutil.move(self.name_file, os.path.join(direct,self.name_file[:-5] + char + self.name_file[-5:]))

class ML(fileXL):
    def __init__(self,name_file):
        fileXL.__init__(self,name_file)
        fil = xlrd.open_workbook(self.name_file)
        sheet = fil.sheet_by_index(0)
        try:
            self.device = sheet.row_values(6)[41]
            self.firm = sheet.row_values(8)[41]
            self.device_name = sheet.row_values(7)[41]
            self.ingerer = sheet.row_values(16)[48]
            self.quantity = int(sheet.row_values(9)[41]) if not sheet.row_values(4)[50] else int (sheet.row_values(11)[41])
            var = sheet.row_values(16)[43]
            if var != '':
                year, month, day = xlrd.xldate_as_tuple(var,0)[:3]
                self.data_vk_end = datetime.date(year, month, day)
            else:
                self.data_vk_end = ''
        except:
            from tkinter import messagebox
            messagebox.showinfo('Я охуел!!!', 'С файлом ' + name_file +' что то не так.....')
            self.device = 'ОШИБКА!!!'
            self.firm = ''
            self.device_name = 'ОШИБКА!!!'
            self.ingerer = ''
            self.quantity = ''
            var = ''

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
            while len(listb.get(index)) != 8:
                index -= 1
            dog = listb.get(index)
            file = (glob(dog + '*' + nomber + '*'))
            if str(file[0][-6]) == '0':
                pril = fileXL(file[0])
                pril.move(GPK_VK,'_В')
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
    else:     #!!!!!!!!!!!!!!!!!!!
        os.chdir(r'\\192.168.1.78\личные папки\12_Апальков Дмитрий Сергеевич\!!!!Новые Договора')
        dog = glob(value[0:6] + '*')
        if dog:
            os.chdir(r'\\192.168.1.78\личные папки\12_Апальков Дмитрий Сергеевич\!!!!Новые Договора' + '\\' + dog[0])
        else:
            dog = glob(value[0:5] + '*')
            os.chdir(r'\\192.168.1.78\личные папки\12_Апальков Дмитрий Сергеевич\!!!!Новые Договора' + '\\' + dog[0])
        for f in os.listdir():
            if 'ПМ' in f and value[0:4] in f:
                os.startfile(f)
                break
        else:
            from tkinter import messagebox
            messagebox.showinfo('Нет методики', 'Программа и методика не обнаружена')
            

def device():
    os.chdir(r'\\192.168.1.78\личные папки\12_Апальков Дмитрий Сергеевич\!!!!Новые Договора\Метрология')
    dev = sorted(glob('EQBASE*.xlsx'))
    os.startfile(dev[-1])
    
def dog_tec():
    os.chdir(r'\\192.168.1.78\личные папки\12_Апальков Дмитрий Сергеевич\!!!!Новые Договора')
    dev = sorted(glob('План выполнение позиций.xlsx'))
    os.startfile(dev[-1])
    
def open_pas_dog():
    os.chdir(r'\\192.168.1.78\личные папки\12_Апальков Дмитрий Сергеевич\!!!!Новые Договора\Акт_изготовления_оснастки')
    dev = sorted(glob('Технологическая оснастка по договору упрощенная.docx'))
    os.startfile(dev[-1])
    
def open_pas_lab():
    os.chdir(r'\\192.168.1.78\личные папки\12_Апальков Дмитрий Сергеевич\!!!!Новые Договора\Акт_изготовления_оснастки')
    dev = sorted(glob('Технологическая оснастка лаборатории упрощенная.docx'))
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
        for i in sorted(glob('*')):
            if re.match(dog, i) is not None:
                pack = i
        os.chdir(os.path.join(DATASHEET,pack))
        res = glob(nomber[:-2] + '*')
        if res:
            os.chdir(os.path.join(DATASHEET,pack,res[0]))
            for f in glob('?' + '*.pdf'):
                os.startfile(f)
        else:
            res = glob(nomber[1:-2] + '*')
            if res:
                os.chdir(os.path.join(DATASHEET,pack,res[0]))
                for f in glob('?' + '*.pdf'):
                    os.startfile(f)
    else:
        os.chdir(r'\\192.168.1.78\личные папки\12_Апальков Дмитрий Сергеевич\!!!!Новые Договора\КГВСЕ')
        for f in os.listdir():
            if 'КГ' in f and value[0:6] in f:
                os.startfile(f)

TODAY =  'Deep Pink'
YESTERDAY = 'OliveDrab'
TREEDAYS =  'green'
TREE_SEVEN =  'indigo'
SEVEN_14 = 'DarkSlateGray'
MORE_14 = 'DarkRed'
           
def get_vk():
    listboxVK.delete(0, END)
    LIST_VK_OBJ = []
    os.chdir(VK)
    global VK_LIST
    VK_LIST = []
    file = sorted(glob('*_.xlsm'))
    for t,f in enumerate(file):
        if f[0] != '~':
            obj = ML(f)
            LIST_VK_OBJ.append(obj)
    number_dog = ''
    for f in LIST_VK_OBJ:
        if f.nomber_dog != number_dog:
            number_dog = f.nomber_dog
            VK_LIST.append(str(t)+ "&&" + f.nomber_dog)
            listboxVK.insert(END, number_dog)
            listboxVK.itemconfig(listboxVK.size() - 1 , foreground='blue')
        res = '  ' + f.nomber_position + ' - ' + f.device + ' - ' + f.firm + ' - ' + f.device_name
        if f.data_vk_end == '':                                                                           ### !!!! Условие красного цвета
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground=MORE_14)
        elif datetime.date.today() - f.data_vk_end == datetime.timedelta(0):
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground=TODAY)
        elif datetime.date.today() - f.data_vk_end == datetime.timedelta(1):     
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground=YESTERDAY)
        elif datetime.timedelta(3) >= datetime.date.today() - f.data_vk_end > datetime.timedelta(1):     
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground=TREEDAYS)
        elif datetime.timedelta(7) >= datetime.date.today() - f.data_vk_end > datetime.timedelta(3):     ### !!!! Условие цвета Индиго
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground=TREE_SEVEN)
        elif datetime.timedelta(14) >= datetime.date.today() - f.data_vk_end > datetime.timedelta(7):     ### !!!! Условие цвета Индиго
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground=SEVEN_14)
        elif datetime.date.today() - f.data_vk_end >= datetime.timedelta(14):
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1, foreground=MORE_14) ### !!!! Условие красного цвета
        else:                             
            listboxVK.insert(END, res)
            listboxVK.itemconfig(listboxVK.size() - 1)
        VK_LIST.append(str(t)+ "&&" + res)
        


def get(lis):
    LIST_OBJ = []
    num = LIST_ING_PFK_D[ComboboxPFK.get()]
    work = Combobox_work.get()
    t = 'р\\d.+р\\d(_ВС|_В_|_ВC_|_В|В|С|С_)?(.*)\\.xls[xm]$'
    if lis == 'pfk':
        listboxPFK.delete(0, END)
        listb = listboxPFK
        direct = GPK_VK
        os.chdir(direct)
        if work == 'В работе':
            file = sorted(glob('*МЛ*' + num + 'р' + '?' + '.xlsm')) + sorted(glob('*МЛ*' + num + 'р' + '?' + '_В.xlsm'))
        if work == 'Работа приостановлена':
            file = sorted(glob('*МЛ*' + num + 'р' + '*' + '.xlsm'))
    elif lis == 'ci':
        listboxCI.delete(0, END)
        listb = listboxCI
        direct = GPK_CI
        os.chdir(direct)
        if work == 'В работе':
            file = sorted(glob('*МЛ*' + num + 'р' + '?' + '_В.xlsm')) + sorted(glob('*МЛ*' + num + 'р' + '?' + '_ВС.xlsm'))+sorted(glob('*МЛ*' + num + 'р' + '?' + '_ВC.xlsm'))
        if work == 'Работа приостановлена':
            file = sorted(glob('*МЛ*' + num + 'р' + '?' + '_В' + '*' + '.xlsm')) + sorted(glob('*МЛ*' + num + 'р' + '?' + '_ВС' + '*' + '.xlsm'))
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
        if work == 'В работе':
            if f.nomber_dog != number_dog:
                number_dog = f.nomber_dog
                listb.insert(END, number_dog)
                listb.itemconfig(listb.size() - 1 , foreground='blue')
            res = '  ' + f.nomber_position + ' - ' + f.device + ' - ' + f.firm + ' - ' + f.device_name
            listb.insert(END, res)
        if work == 'Работа приостановлена' and lis not in('Lesha','nky','nky2'):
            if list(re.findall(t, f.name_file))[0][-1] != '':
                if f.nomber_dog != number_dog:
                    number_dog = f.nomber_dog
                    listb.insert(END, number_dog)
                    listb.itemconfig(listb.size() - 1 , foreground='blue')
                res = '  ' + f.nomber_position + ' - ' + f.device + ' - ' + f.firm + ' - ' + f.device_name + ' - ' + list(re.findall(t, f.name_file))[0][-1]
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
    value2 = w.get(index)
    if len(value) != 8:
        i = index
        while i >= 0 and len(value) != 8:
            value = w.get(i)
            i-=1
        i = value
        x = '-'.join([value, re.findall('\d{4}\.\d', value2)[0]])
        os.chdir(VK)
        print(x)
        file = (glob(x + '*'))
        print(file)
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
    check()
    
def finish():
    if LEL_DOG != '':
        from tkinter import messagebox
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
        from tkinter import messagebox
        if messagebox.askyesno('Перенос', 'Вы действительно хотите перенести позицию ' + LEL_DOG  + '-' + LEL_NOMBER  + ' на следующий этап?'):
            messagebox.showinfo('Да', 'Хорошо')
            os.chdir(GPK_VK)
            if len(glob(LEL_DOG + '-' + LEL_NOMBER + '*ПЛ*' + '*')) == 1:
                for pl in sorted(glob( LEL_DOG + '-' + LEL_NOMBER + '*ПЛ*' + '*')):
                    pril = fileXL(pl)
                    pril.copy(GPK_VK_O)
                    print(pril.name_file,'перенесен в отчетные')
                    pril.move(GPK_CI,'С')
                    print(pril.name_file,'перенесен в CИ')
            if len(glob( LEL_DOG + '-' + LEL_NOMBER + '*МЛ*' + '*')) == 1:
                for ml in sorted(glob( LEL_DOG + '-' + LEL_NOMBER + '*МЛ*' + '*')):
                    marsh = ML(ml)
                    marsh.copy(GPK_VK_O)
                    print(marsh.name_file,'перенесен в отчетные')
                    marsh.copy(CKK)
                    print(marsh.name_file,'перенесена в ССК')
                    marsh.move(GPK_CI,'С')
                    print(marsh.name_file,'перенесен в СИ')
            get('pfk')
            get('ci')
            messagebox.showinfo('Перенос', 'Позиция перенесена =)')
        else:
            messagebox.showinfo('Неа', 'Ну и ладно')

def finish_ci():
    if LEL_DOG != '':
        from tkinter import messagebox
        if messagebox.askyesno('Перенос', 'Вы действительно хотите перенести позицию ' + LEL_DOG  + '-' + LEL_NOMBER  + ' на следующий этап?'):
            messagebox.showinfo('Да', 'Хорошо')
            os.chdir(GPK_CI)
            if len(glob(LEL_DOG + '-' + LEL_NOMBER + '*ПЛ*' + '*')) == 1:
                for pl in sorted(glob( LEL_DOG + '-' + LEL_NOMBER + '*ПЛ*' + '*')):
                    marsh = ML(pl)
                    marsh.move(GPK_CI_O)
                    print(pril.name_file,'перенесена в отчетные')
            if len(glob( LEL_DOG + '-' + LEL_NOMBER + '*МЛ*' + '*')) == 1:
                for ml in sorted(glob( LEL_DOG + '-' + LEL_NOMBER + '*МЛ*' + '*')):
                    marsh = ML(ml)
                    marsh.copy(GPK_VK_O)
                    print(marsh.name_file,'перенесен в отчетные')
                    marsh.move(GPK_MKI)
                    print(marsh.name_file,'перенесен к Алексею')
            get('ci')
            get('Lesha')
            messagebox.showinfo('Перенос', 'Позиция перенесена =)')
        else:
            messagebox.showinfo('Неа', 'Ну и ладно')

def repeat():
    while (var1.get()):
        time.sleep(900)
        get_vk()
        get('nky')
        

def check():
    t1 = threading.Thread(target=repeat)
    if  var1.get() == 1:
        t1.start()
                              
if  __name__ ==  "__main__" :

    # Папки и катологи
    VK = r'\\192.168.1.78\database\05-Отчетные\03-01_ГВК МЛ'
    GPK_VK = r'\\192.168.1.78\database\05-Отчетные\03-03_ГПК МЛ+ПЛ+ПС НКУ'
    GPK_VK_O = r'\\192.168.1.78\database\05-Отчетные\03-04_ГПК МЛ+ПЛ+ПС НКУ  О'
    CKK = r'\\192.168.1.78\database\05-Отчетные\03-05_СКК МЛ'
    GPK_CI = r'\\192.168.1.78\database\05-Отчетные\04-01_ГПК МЛ+ПЛ+ПС МКИ'
    GPK_CI_O = r'\\192.168.1.78\database\05-Отчетные\04-02_ГПК МЛ+ПЛ+ПС МКИ О'
    GPK_MKI = r'\\192.168.1.78\database\05-Отчетные\04-03_ГМК МЛ'
    NKY = r'\\192.168.1.78\database\05-Отчетные\04-05_ГВК МЛ МКИ'
    NKY2 = r'\\192.168.1.78\database\05-Отчетные\05-01_ГПК МЛ+ПЛ+ПС НКУ'
    NKY2_O = r'\\192.168.1.78\database\05-Отчетные\05-02_ГПК МЛ+ПЛ+ПС НКУ  О'
    CKK_2 = r'\\192.168.1.78\database\05-Отчетные\05-03_СКК МЛ'
    DATASHEET = r'\\192.168.1.78\datasheet\02_Текущие работы'
    # Список сотрудников
    LIST_ING_PFK_D = {'Апальков Д.С.': '012','Ломазов Е.Е.': '022','Данилкин П. П.': '041','Михайлюченко И. В.': '045','Шершеньков М. Б.': '049','Тупиков П. А.': '058','Лапидус П.К.': '059','Ларионов В. В.': '071','Крохалев И. Н.':'073','Третьяков С. А.':'080','ВСЕ':'*'}
    # Выбранный элемент
    LEL_DOG = ''
    LEL_NOMBER = ''
    root = Tk()
    global var1
    var1=IntVar()
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
    listboxVK = Listbox(root, height=17, width=60, selectmode=EXTENDED)
    listboxPFK = Listbox(root, height=17, width=60, selectmode=EXTENDED)
    ComboboxPFK = Combobox(root, values = list(sorted(LIST_ING_PFK_D.keys())), height=len(list(sorted(LIST_ING_PFK_D.keys()))))
    Combobox_work = Combobox(root, values = ['В работе','Работа приостановлена'], height=1 )
    ComboboxPFK.set('Апальков Д.С.')
    Combobox_work.set('В работе')
    listboxCI = Listbox(root, height=17, width=60, selectmode=EXTENDED)
    listboxLesha = Listbox(root, height=17, width=60, selectmode=EXTENDED)
    listboxNKY = Listbox(root, height=17, width=60, selectmode=EXTENDED)
    listboxNKY2 = Listbox(root, height=17, width=60, selectmode=EXTENDED)
    rossupT = Text(root,height=0.5,width=12,font='Arial 14',wrap=WORD)
    rossupL = Label(root, text = "Количество:")
    Text_time1 = Label(root, text = "Позиции текущие", fg = TODAY,font='Arial 15')
    Text_time2 = Label(root, text = "1 день", fg = YESTERDAY,font='Arial 15')
    Text_time3 = Label(root, text = "2-3 дня", fg = TREEDAYS,font='Arial 15')
    Text_time4 = Label(root, text = "3-7 дней", fg = TREE_SEVEN,font='Arial 15')
    Text_time5 = Label(root, text = "7-14 дней", fg = SEVEN_14,font='Arial 15')
    Text_time6 = Label(root, text = ">14 дней", fg = MORE_14,font='Arial 15')
    vk_text = Label(root, text = "Готовность входного контроля:",font='Arial 14')
    pfk_text = Label(root, text = "ВК ПФК в работе:",font='Arial 14')
    ci_text = Label(root, text = "СИ ПФК в работе:",font='Arial 14')
    Lesha_text = Label(root, text = "МКИ в работе:",font='Arial 14')
    NKY_text = Label(root, text = "Готовность после МКИ:",font='Arial 14')
    NKY2_text = Label(root, text = "СИ НКУ в работе:",font='Arial 14')
    finishVK_T = Text(root,height=0.5,width=12,font='Arial 14',wrap=WORD)
    finishVK_L = Label(root, text = "Окончен ВК:")
    engineer_VK_T = Text(root,height=0.5,width=12,font='Arial 14',wrap=WORD)
    engineer_VK_L = Label(root, text = "Инженер ВК:")
    device = Button(root, text='Приборы',width=10,bg='black',fg='red', command = device)
    dog_tec = Button(root, text='ПЛАН ДИМЫ',width=15,bg='blue',fg='yellow', command = dog_tec)
    pasport = Button(root, text='Форма паспорта о договору',width=23,bg='black',fg='green', command = open_pas_dog)
    pasport_lab = Button(root, text='Форма паспорта лаборатории',width=23,bg='brown',fg='black', command = open_pas_lab)
    buttonVK_CI = Button(root, text = '>>', command=finish_vk)
    buttonCI_Lesha = Button(root, text = '>>', command=finish_ci)
    buttonCI_Finish = Button(root, text = '>>', command = finish)
    auto_updates = Checkbutton(root,text=u'Автообновление',variable=var1 ,onvalue=1 ,offvalue=0,command=check)
    update()


    buttonVK.place(x=110, y=0) # Игорёк))))
    listboxVK.place(x=0, y=140)
    listboxVK.bind('<<ListboxSelect>>',sel_vk)
    listboxVK.bind('<Double-Button-1>',open_file)
    listboxVK.bind('<Button-3>',datasheet)
    listboxPFK.place(x=400, y=140)
    listboxPFK.bind('<<ListboxSelect>>',selected)
    ComboboxPFK.bind('<<ComboboxSelected>>', update_ing)
    listboxPFK.bind('<Double-Button-1>',open_file)
    listboxPFK.bind('<Button-3>',datasheet)
    listboxCI.place(x=800, y=140)
    listboxCI.bind('<<ListboxSelect>>',selected)
    listboxCI.bind('<Double-Button-1>',open_file)
    listboxCI.bind('<Button-3>',datasheet)
    listboxLesha.place(x=1200, y=140)
    listboxLesha.bind('<Double-Button-1>',open_file)
    listboxLesha.bind('<Button-3>',datasheet)
    listboxNKY.place(x=0, y=660)
    listboxNKY.bind('<Double-Button-1>',open_file)
    listboxNKY.bind('<Button-3>',datasheet)
    listboxNKY2.place(x=400, y=660)
    listboxNKY2.bind('<<ListboxSelect>>',selected)
    listboxNKY2.bind('<Double-Button-1>',open_file)
    listboxNKY2.bind('<Button-3>',datasheet)
    vk_text.place(x=51, y=105)
    pfk_text.place(x=515, y=105)
    ci_text.place(x=920, y=105)
    Lesha_text.place(x=1330, y=105)
    NKY_text.place(x=70, y=620)
    NKY2_text.place(x=510, y=620)
    auto_updates.place(x=500, y = 470)
    rossupT.place(x=100, y = 420)
    rossupL.place(x=0, y = 420)
    finishVK_T.place(x=100, y = 450)
    finishVK_L.place(x=0, y = 450)
    engineer_VK_T.place(x=100, y = 480)
    engineer_VK_L.place(x=0, y = 480)
    device.place(x=250, y = 420)
    dog_tec.place(x=250, y = 450)
    pasport.place(x=250, y = 480)
    pasport_lab.place(x=250, y = 510)
    LabelPFK.place(x=530, y = 0)
    ComboboxPFK.place(x=500, y = 450)
    Combobox_work.place(x=500, y = 430)
    buttonVK_CI.place(x=770, y = 270)
    LabelCI.place(x=930, y = 0)
    LabelLesha.place(x=1330, y = 0)
    buttonCI_Lesha.place(x=1170, y = 270)
    LabelNKY.place(x=110, y=520)
    LabelNKY2.place(x=530, y=520)
    LabelFinish.place(x=800, y=640)
    buttonCI_Finish.place(x=770, y = 770)
    Text_time1.place(x=800, y = 410)
    Text_time2.place(x=800, y = 450)
    Text_time3.place(x=800, y = 490)
    Text_time4.place(x=800, y = 530)
    Text_time5.place(x=800, y = 570)
    Text_time6.place(x=800, y = 610)
    root.mainloop()
