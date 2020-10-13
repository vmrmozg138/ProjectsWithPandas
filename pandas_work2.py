import pandas as pd
from tkinter import filedialog as tkd

fPathStr = tkd.askopenfilename()
fname = fPathStr[fPathStr.rfind('/'):fPathStr.rfind('.')]
print(fPathStr,'\n',fname)

if fPathStr.endswith('xlsx'):
    data = pd.read_excel(fPathStr)
elif fPathStr.endswith('csv'):
    data = pd.read_csv(fPathStr,encoding='utf-16',delimiter='\t')

df = pd.DataFrame(data)

mmColl = {}

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def toNormalview(inp):
    if is_number(inp):
        return str(int(inp))
    else:
        return inp

for index,row in df.iterrows():
    if toNormalview(row[1]) not in mmColl:
        mmColl[toNormalview(row[1])] = []
    mmColl[toNormalview(row[1])].append(toNormalview(row[3]))
    

for k,v in mmColl.items():
    print(k,v)

with open(fPathStr[:fPathStr.rfind('/')+1]+fname+".txt", "w", encoding="utf-8") as f:
    for k,v in mmColl.items():
        #это для марок и моделей
        output = '\t\tcase {_'+str(k)+'}\n\t\t\tmodelFilter = {'+','.join('_'+str(element) for element in v)+'}\n'     
        #для сегментов
        #output = '\t\tif len(thatExactCar*segment'+ (str(k)).replace(' ','_') +')>0 then\n\t\t\tsegmentFilter = {'+ ','.join('_'+str(element) for element in v)+'}\n\t\tend if\n'
        #для регионов
        #output = '\t\tif SQ3*{'+ ','.join('_'+str(element) for element in v) +'} then\n\t\t\tSQ3a = {_'+ (str(k)).replace(' ','_') +'}\n'
        f.write(output)


