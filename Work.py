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

root = Tk()
root.title('Урок 2')
root.iconbitmap('alfa-komplekt-logo-1458826805.ico')
root.geometry('1200x600')
lab1 = Label(root, text="ВК", font="Calibri 14")
listboxVK = Listbox(root, height=50, width=50, selectmode=EXTENDED)
os.chdir(VK)
file = sorted(glob('*_.xlsm'))
number_dog = []
for f in file:
    if f[0:8] != number_dog and f[0] != '~' and f[0:8] != '149-01-1':
        number_dog = f[0:8]
        listboxVK.insert(END, number_dog)
    if f[0] != '~' and f[0:8] != '149-01-1':
        fil = xlrd.open_workbook(f)
        sheet = fil.sheet_by_index(0)
        val = sheet.row_values(6)[9]
        pos = '  ' + f[9:15] + ' - ' + val
        listboxVK.insert(END, pos)

lab1.place(x=130, y=0)
listboxVK.place(x=0, y=25)
root.mainloop()
