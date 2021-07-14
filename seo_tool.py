from urllib.request import urlopen
from bs4 import BeautifulSoup
import urllib.request
import re
import string
import xlsxwriter
url="https://www.tutorialspoint.com//python/index.htm"
#"https://www.tutorialspoint.com//python/index.htm"
req=urllib.request.Request(url)
f=urllib.request.urlopen(req)
html=f.read().decode('utf-8')
soup=BeautifulSoup(html,"html.parser")
for script in soup(["script","style"]):
    script.extract()
frequency = {}
text=soup.get_text()
lines=(line.strip() for line in text.splitlines())
for line in lines:
    line=line.lower()
    j=re.sub(' +',' ' ,line)
    text_string = j
    match_pattern = re.findall(r'\b[a-z|A-Z]{3,15}\b', text_string)
    for word in match_pattern:
        count = frequency.get(word,0)
        frequency[word] = count + 1
        #print(frequency[word])
frequency_list = frequency.keys()
workbook = xlsxwriter.Workbook('prjchrt.xlsx')
worksheet = workbook.add_worksheet()
# Create a new Chart object.
chart = workbook.add_chart({'type': 'column'})
ss=[]
x=0
for words in frequency_list:
    if(frequency[words]>5):
        ss.append([])
        print(words, frequency[words])
        ss[x].append(words)
        ss[x].append(frequency[words])
        x += 1
def takeSecond(elem):
    return elem[1]
ss.sort(key=takeSecond)
print("\n\n ---------------------Ordered List------------------------ \n\n")
print(ss)
#print(type(ss))
n=len(ss)
#print(n)

    
for col,rdata in enumerate(ss):
    for row,cdata in enumerate(rdata):
        worksheet.write(row,col,cdata)
#worksheet.write_column('A1', ss[0])
#worksheet.write_column('B1', ss[1])
#worksheet.write_column('C1', ss[2])
#worksheet.write_column('D1', ss[3])
#worksheet.write_column('E1', ss[4])
#worksheet.write_column('F1', ss[5])
#worksheet.write_column('G1', ss[6])
#worksheet.write_column('H1', ss[7])
#worksheet.write_column('I1', ss[8])
# Configure the chart. In simplest case we add one or more data series.
chart.add_series({'values': '=Sheet1!$A$1:$A$5'})
chart.add_series({'values': '=Sheet1!$A$:$B$5'})
chart.add_series({'values': '=Sheet1!$C$1:$C$5'})
chart.add_series({'values': '=Sheet1!$D$1:$D$5'})
chart.add_series({'values': '=Sheet1!$E$1:$E$5'})
chart.add_series({'values': '=Sheet1!$F$1:$F$5'})
chart.add_series({'values': '=Sheet1!$G$1:$G$5'})
chart.add_series({'values': '=Sheet1!$H$1:$H$5'})
chart.add_series({'values': '=Sheet1!$I$1:$I$5'})
chart.add_series({'values': '=Sheet1!$J$1:$J$5'})
chart.add_series({'values': '=Sheet1!$K$1:$K$5'})
chart.add_series({'values': '=Sheet1!$L$1:$L$5'})
chart.add_series({'values': '=Sheet1!$M$1:$M$5'})
chart.add_series({'values': '=Sheet1!$N$1:$N$5'})
chart.add_series({'values': '=Sheet1!$O$1:$O$5'})
chart.add_series({'values': '=Sheet1!$P$1:$P$5'})
chart.add_series({'values': '=Sheet1!$Q$1:$Q$5'})
chart.add_series({'values': '=Sheet1!$R$1:$R$5'})
chart.add_series({'values': '=Sheet1!$S$1:$S$5'})
chart.add_series({'values': '=Sheet1!$T$1:$T$5'})
chart.add_series({'values': '=Sheet1!$U$1:$U$5'})
chart.add_series({'values': '=Sheet1!$V$1:$V$5'})
chart.add_series({'values': '=Sheet1!$W$1:$W$5'})
chart.add_series({'values': '=Sheet1!$X$1:$X$5'})
chart.add_series({'values': '=Sheet1!$Y$1:$Y$5'})
chart.add_series({'values': '=Sheet1!$Z$1:$Z$5'})

worksheet.insert_chart('B7', chart)
workbook.close()
