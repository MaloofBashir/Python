import pandas as pd 
from docx import Document

HST322J=[]

import sys 

def extract_table(table_index,user_name):
    document=Document(r"colg.docx")

    table_data=[]
    table=document.tables[table_index]
    for row in table.rows:
        row_data=[]
        for cell in row.cells:
            row_data.append(cell.text.strip())
        table_data.append(row_data)
    df=pd.DataFrame(table_data[1:],columns=table_data[0])

    new_count=0

    for index,row in df.iterrows():
        value=row["Subjects1"]
        if user_name in value:
            value=df.loc[index,"ClassRollNo"]
            if value.isdigit()!=True:
                print(df.iloc[index])
                new_count+=1
            HST322J.append(value)

for i in range(17):
    extract_table(i,"PLS322J")

print(HST322J)
print(f"total no of studetns are {len(HST322J)}")
count=0

with open("roolno.txt",'w') as file:
    for data in HST322J:
        file.write(data+ ",")


for i in HST322J:
    if i.isdigit():
        count=count+1
        
print(f"new counter value is {count}")
