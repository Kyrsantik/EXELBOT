#import telebot
#import openpyxl
from openpyxl import load_workbook


booking = load_workbook("11_11.xlsx")
sheeting = booking["ГРУППЫ"]
"""
y =0
x= 0
#y = y + 1
#x = "A"+ str(y)
for i in range(3,10):
    y = y + 1
    x = "A" + str(y)
    print(sheeting[str(x)].value)

booking.close()
"""
for row in range(1,34):
    c = sheeting[row][0].value
    b = sheeting[row][1].value
    d = sheeting[row][2].value
    e = sheeting[row][3].value
    f = sheeting[row][4].value
    g = sheeting[row][5].value
    h = sheeting[row][6].value
    p = sheeting[row][7].value
    q = sheeting[row][8].value
    r = sheeting[row][9].value
    s = sheeting[row][10].value
    t = sheeting[row][11].value
    x = sheeting[row][12].value
    if [] == "None":
        [] == ""
    print(c,b,d,e,f,g,h,p,q,r,s,t,x if [] == "None" else c,b,d,e,f,g,h,p,q,r,s,t,x )



