__author__ = 'Тупиков Павел'
__version__ = 1.0

VK = r'F:\Работа\03-01_ГВК МЛ'
GPK_VK = r'F:\Работа\03-03_ГПК МЛ+ПЛ+ПС НКУ'
GPK_VK_O = r'F:\Работа\03-04_ГПК МЛ+ПЛ+ПС НКУ  О'
CKK = r'F:\Работа\03-05_СКК МЛ'
GPK_CI_O = r'F:\Работа\04-02_ГПК МЛ+ПЛ+ПС МКИ О'
#D04-03_ГМК МЛ = r'F:\Работа\04-03_ГМК МЛ'
#D04-05_ГВК МЛ МКИ = r'F:\Работа\04-05_ГВК МЛ МКИ'
GPK_MKI = r'F:\Работа\05-01_ГПК МЛ+ПЛ+ПС НКУ'
GPK_MKI_O = r'F:\Работа\05-02_ГПК МЛ+ПЛ+ПС НКУ  О'
#D05-03_СКК МЛ = r'F:\Работа\05-03_СКК МЛ'

from tkinter import *
import os
from glob import glob
import xlrd
from tkinter import *


LIST_VK =[]


def get_vk():
    os.chdir(VK)
    file = sorted(glob('*_.xlsm'))
    number_dog = []
    for f in file:
        if f[0:8] != number_dog and f[0] != '~' and f[0:8] != '149-01-1':
            number_dog = f[0:8]
            listboxVK.insert(END, number_dog)
            LIST_VK.append(number_dog)
        if f[0] != '~' and f[0:8] != '149-01-1':
            fil = xlrd.open_workbook(f)
            sheet = fil.sheet_by_index(0)
            val = sheet.row_values(6)[41]
            val2 = sheet.row_values(8)[41]
            val3 = sheet.row_values(7)[41]
            pos = '  ' + f[9:15] + ' - ' + val + ' - ' + val2 + ' - ' + val3
            listboxVK.insert(END, pos)
            LIST_VK.append(pos)

def select(event):
    w = event.widget
    index = int(w.curselection()[0])
    value = w.get(index)
    print ('You selected item %d: "%s"' % (index, value))

    
def select_but():
    listboxVK.delete(0, END)
    get_vk()

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


root = Tk()
root.title('РАБОТА')
root.iconbitmap('alfa-komplekt-logo-1458826805.ico')
root.geometry('800x600')
image = PhotoImage(file = 'ВК.png')
buttonVK = Button(root, image=image, command=select_but)
listboxVK = Listbox(root, height=15, width=60, selectmode=EXTENDED)
rossupE = Entry(root, width=10, bd=3)
rossupL = Label(root, text = "Лента/Росыпь:")
finishVK_E = Entry(root, width=10, bd=3)
finishVK_L = Label(root, text = "Окончен ВК:")
get_vk()


buttonVK.place(x=80, y=0) # Игорёк))))
listboxVK.place(x=0, y=175)
listboxVK.bind('<<ListboxSelect>>',select)
listboxVK.bind('<Double-Button-1>',open_file)
rossupE.place(x=100, y = 420)
rossupL.place(x=0, y = 420)
finishVK_E.place(x=250, y = 420)
finishVK_L.place(x=170, y = 420)
root.mainloop()
