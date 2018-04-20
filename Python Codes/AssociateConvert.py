import xlrd
import xlwt

rb = xlrd.open_workbook('DataSet.xls')
rs = rb.sheet_by_index(0)


wb = xlwt.Workbook()
ws = wb.add_sheet('DataSet')
ws.write(0,0,"ResourceChange")
ws.write(0,1,"EconomyChange")
ws.write(0,2,"FleetChange")
ws.write(0,3,"ResearchChange")

i=1
while(i>=1):
    try:
        for x in range(4):
            if(rs.cell_value(i,3+x*3)[0:1]=="+"):
                ws.write(i,x,"1")
            elif(rs.cell_value(i,3+x*3)[0:1]=="-"):
                ws.write(i,x,"2")
            else:
                ws.write(i,x,"0")
        i+=1
    except:
        i=-1

tableName="Assoicate"
print "Saving Excel as "+tableName+".xls"
wb.save(tableName+".xls")
