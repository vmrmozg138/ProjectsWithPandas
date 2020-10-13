import pandas as pd
import csv
import xlrd
import re 
from tkinter import filedialog as tkd

file_path_string = tkd.askopenfilename()

lines_length = []

with open(file_path_string, 'r',encoding = 'utf-8') as csvfile:
    csv_reader = csv.reader(csvfile, dialect='excel', delimiter = ';')
    lreader = list(csv_reader)
    header = lreader[0]
'''    for i, line in enumerate(lreader):
        a = len(line)
        if i == 0:
            header = line
        lines_length.append(a) 
'''


print(len(header))
corrupted = {}

for i, elem in enumerate(header):
    if re.search("/",elem):
        base = elem.split(' ')[0]
        if base not in corrupted:
            corrupted[base] = []
        corrupted[base].append(i)


print('\n\n',corrupted)
