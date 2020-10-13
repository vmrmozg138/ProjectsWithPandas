from tkinter import filedialog as tkd
import pandas as pd
import numpy as np
import chardet

file_path_string = tkd.askopenfilename()
with open(file_path_string, 'rb') as f:
    resutl = chardet.detect(f.read())

print(resutl['encoding'])

csv_obj = pd.read_csv(file_path_string,encoding=resutl['encoding'],delimiter='\t',index_col='Respondent.Serial')
df = pd.DataFrame(csv_obj)

qtcols = [col for col in df.columns if "QuestionTimers" in col]

print(qtcols[0])

finalColSet = ['ID','ReturnCode']
finalColSet.append(str(qtcols[0]))
print(finalColSet)

df_processed = df[finalColSet].loc[df['ReturnCode']=='C']
df_processed1 = df[str(qtcols[0])].loc[df['ReturnCode']=='C']
print(df_processed)

df_processed.to_csv(file_path_string[:file_path_string.rfind('/')+1]+"output.csv",sep='\t')
df_processed1.to_csv(file_path_string[:file_path_string.rfind('/')+1]+"output1.txt",sep='\t')

#extracolset = ''+



