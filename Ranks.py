import xlwt 
wb = xlwt.Workbook()

style = xlwt.XFStyle()
font = xlwt.Font()
font.bold = True
style.font = font

def k(dct): #Returns name with maximum GP
    m = max(dct.values())
    for i in dct.keys():
        if dct[i] == m:
            return i

nfile = open("Names_CSE.txt","r") #Change filename here for names.
names = {}
mx = 0
while 1:
    y = nfile.readline().split()
    if not y:
        break
    #y.pop()
    #y.pop()
    nm = ''
    for i in range(2,len(y)):
        nm += y[i] + " "
    names[y[1]] = nm
    if len(nm)>mx:
        mx = len(nm)

ifile = open("Grades_CSE.txt",'r') #Change filename here for grades.
dct = {}
dct_dept = {"BT":"Biotech","CH":"Chemical","EP":"E.Physics","CE":"Civil","CS":"Comp.Sc.","EE":"EEE","EC":"ECE","ME":"Mech","PE":"Production"}
cnt = 0
while 1:
    y = ifile.readline().split()
    if not y:
        break
    dct[y[1]] = float(y[-1]) #y[1] is Roll number while y[-1] is the Grade
    if(y[-1][-1].isnumeric()==False):
        print(y[1],"CGPA Error\n")
    cnt+=1
ifile.close()
sheet1 = wb.add_sheet("Rankings")
sheet1.col(0).width = 1000
sheet1.col(1).width = 16000
sheet1.col(2).width = 6000
sheet1.col(3).width = 6000
sheet1.col(4).width = 2000
sheet1.write(0,0,'No.',style = style)
sheet1.write(0,1,'Name',style = style)
sheet1.write(0,2,'Roll no.',style = style)
sheet1.write(0,3,'Department',style = style)
sheet1.write(0,4,'SGPA',style = style)
for i in range(1,cnt+1):
    sheet1.write(i,0,str(i))
    sheet1.write(i,1,names[k(dct)])
    sheet1.write(i,2,k(dct))
    sheet1.write(i,3,dct_dept[k(dct)[7:]])
    sheet1.write(i,4,str(dct[k(dct)]))
    del dct[k(dct)] #Deletes the top most name
wb.save('RankingsCSE.xls') #Output file name

