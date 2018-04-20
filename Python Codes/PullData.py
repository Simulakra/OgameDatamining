import urllib2
from HTMLParser import HTMLParser

global inTable, inData, ofHead
global count, longData, tableData
inTable=0
ofHead=0
inRow=0
count=0
tableData=[]
rowData=[]

class MyHTMLParser(HTMLParser):
    def handle_starttag(self, tag, attrs):
        global inTable, inRow, ofHead
        if(tag == "table"):
            inTable=1
        if(tag == "tr"):
            inRow=1
        if(tag == "thead"):
            ofHead=0
        
    def handle_endtag(self, tag):
        global inTable, inRow, ofHead
        global count, rowData,tableData
        if(tag == "table"):
            inTable=0
        if(tag == "tr"):
            inRow=0
            if(ofHead):
                #print rowData
                tableData.append(rowData)
                rowData=[]
        if(tag == "thead"):
            ofHead=1
       
    def handle_data(self, data):
        global inTable, inRow, ofHead
        global count, rowData
        if(inTable & inRow & ofHead):
            if(len(data.strip())!=0):
                rowData.append(data.strip())
                
                
def getDataFromLink(webpage):
    parser = MyHTMLParser()
    #webpage='http://www.infuza.com/es/tr.ogame.org/Andromeda/Resources/Points/0'
    response = urllib2.urlopen(webpage)
    data = response.read()
    print "*Parsing Website: "+ webpage[38:]
    parser.feed(data)
#-----------------------------------------------------------------------------
universes = ["Universe1"]
universes.append("Universe50")
universes.append("Andromeda")
universes.append("Hydra")
universes.append("Pegasus")
universes.append("Taurus")
universes.append("Ursa")
universes.append("Betelgeuse")
universes.append("Eridanus")
universes.append("JupiterAscending")
universes.append("Ganimed")
universes.append("Hyperion")
universes.append("Izar")
universes.append("Japetus")
universes.append("Callisto")
universes.append("Libra")
universes.append("Merkur")
universes.append("Nusakan")
universes.append("Oberon")
universes.append("Polaris")
universes.append("Quaoar")
universes.append("Rhea")
universes.append("Spica")
universes.append("Tarazed")
universes.append("Uriel")
universes.append("Virgo")
universes.append("Wezn")
universes.append("Xanthus")

import xlwt

wb = xlwt.Workbook()
ws = wb.add_sheet('DataSet')

for x in range(14):
    ws.col(x).width=4325

ws.write(0,0,"Universe")
ws.write(0,1,"Nickname")
ws.write(0,2,"Resource Rank")
ws.write(0,3,"Resource Increase")
ws.write(0,4,"Resource Point")
ws.write(0,5,"Economy Rank")
ws.write(0,6,"Economy Increase")
ws.write(0,7,"Economy Point")
ws.write(0,8,"Fleet Rank")
ws.write(0,9,"Fleet Increase")
ws.write(0,10,"Fleet Point")
ws.write(0,11,"Research Rank")
ws.write(0,12,"Research Increase")
ws.write(0,13,"Research Point")

rowcount=0

#for u_c in range(4):
for u_c in range(len(universes)):   
    colNames=["Resources","Economy","Fleet","Research"]
    finalData=[]
    dataCount=5
    
    for x in range(dataCount):
        getDataFromLink("http://www.infuza.com/es/tr.ogame.org/"+universes[u_c]+"/"+colNames[0]+"/Points/"+str(x))
    
    for x in range(len(tableData)):
        try:
            tableData[x].remove(tableData[x][4])
        except:
            continue
    finalData = tableData
    tableData = []
    for col_index in range(len(colNames)-1):
        global finalData,tableData
        for x in range(dataCount):
            getDataFromLink("http://www.infuza.com/es/tr.ogame.org/"+universes[u_c]+"/"+colNames[col_index+1]+"/Points/"+str(x))
        for x in range(len(finalData)):
            for y in range(len(tableData)):
                if(tableData[y][3]==finalData[x][3]):
                    finalData[x].append(tableData[y][0])
                    finalData[x].append(tableData[y][1])
                    finalData[x].append(tableData[y][2])
                    #print finalData[x]
        tableData = []
    
    print "Writing Table Results of " + universes[u_c]
    for x in range(len(finalData)):
        global rowcount
        finalData[x].insert(0,universes[u_c])
        if(len(finalData[x])==14):
            ws.write(rowcount+1, 0, finalData[x][0])
            ws.write(rowcount+1, 1, finalData[x][4])
            ws.write(rowcount+1, 2, finalData[x][1])
            ws.write(rowcount+1, 3, finalData[x][2])
            ws.write(rowcount+1, 4, finalData[x][3])
            ws.write(rowcount+1, 5, finalData[x][5])
            ws.write(rowcount+1, 6, finalData[x][6])
            ws.write(rowcount+1, 7, finalData[x][7])
            ws.write(rowcount+1, 8, finalData[x][8])
            ws.write(rowcount+1, 9, finalData[x][9])
            ws.write(rowcount+1, 10, finalData[x][10])
            ws.write(rowcount+1, 11, finalData[x][11])
            ws.write(rowcount+1, 12, finalData[x][12])
            ws.write(rowcount+1, 13, finalData[x][13])
            rowcount+=1

print "Saving Excel as DataSet.xls"
wb.save('DataSet.xls')
