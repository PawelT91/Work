import re
import os
import shutil

class fileXL:
    def __init__(self,name_file):
        self.name_file = name_file
        self.nomber_dog = re.findall('\d{3}\-\d{2}\-\d{1}',name_file)[0]
        self.nomber_position = re.findall('\d{4}\.\d{1}', name_file)[0]

    def copy(self, direct, char = ''):
        shutil.copyfile(self.name_file, os.path.join(direct,self.name_file[:-5] + char + self.name_file[-5:]))

    def move(self, direct, char = ''):
        shutil.move(self.name_file, os.path.join(direct,self.name_file[:-5] + char + self.name_file[-5:]))
