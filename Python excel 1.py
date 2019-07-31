# See full Toturial at my Youtube Channel(YB TV): https://www.youtube.com/channel/UCvnhhDKv5takEN412dmVW8g/featured
# GitHab Page:https://github.com/yasser64b/
#Email: big3del@gmail.com


from openpyxl import Workbook # pip install openpyxl 
from openpyxl.styles import Font, Color, colors
 # see example 1 or
listA=['DistanceA (m)',1,2,3,4,5,6,7,8,8,9,9]
listB=['DistanceB (m)',1,2,3,4,5,6,7,8,8,9,9]
L=[listA, listB]
wb = Workbook()
ws1 = wb.active # work sheet
ws1.title = "Pyxl"
for i in range(len(listA)):
    for j in range (len(L)):
        _ = ws1.cell(column=j+1, row=i+1, value=L[j][i])
    # ws1.append([listA[row]])

ft_h = Font(name='Calibri',color=colors.BLUE, bold=True, size=11,underline='double')
ft_b = Font(name='Calibri',color=colors.RED, bold=True, size=11,underline='single')
a1 = ws1['A1']
b1 = ws1['B1']
a1.font = ft_h
b1.font = ft_h

for i in range (2,13):
        a=ws1['A'+str(i)]
        b=ws1['B'+str(i)]
        a.font=ft_b
        b.font=ft_b

ws1["A13"]='=SUM(A2:A12)'
ws1["B13"]='=SUM(B2:B12)'

wb.save('xlfile.xlsx')
