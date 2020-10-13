import pandas as pd
from tkinter import filedialog as tkd

fPathStr = tkd.askopenfilename()
fname = fPathStr[fPathStr.rfind('/')+1:fPathStr.rfind('.')]
print(fPathStr,'\n',fname)

if fPathStr.endswith(('xlsx','xls')):
    data = pd.read_excel(fPathStr)
elif fPathStr.endswith('csv'):
    data = pd.read_csv(fPathStr,encoding='utf-16',delimiter='\t')

df = pd.DataFrame(data)
dfCint = df[df['Supplier']==10168]


print(dfCint['Status'].value_counts())