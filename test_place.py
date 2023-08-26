import openpyxl
work_table = openpyxl.load_workbook(filename='scaler_table.xlsx')
sheet = work_table['Scaler_Place']
vals=[]
i=0
x=''

while x!=None:
    i=i+1
    x=sheet['A'+str(i)].value
    vals.append(x)
print(vals)
for Inspector in vals:
    if Inspector==4:
        print(vals.index(Inspector))
        sheet['A'+str(vals.index(Inspector)+1)]='LOL'
work_table.save('scaler_table.xlsx')