#Required libraries:  Pandas, pdfkit
#Install tool: wkhtmltopdf using 'sudo apt-get install wkhtmltopdf' on linux


import pandas as pd
import warnings
import pdfkit as pdf

warnings.filterwarnings("ignore", category=FutureWarning)

#read excel file
data = pd.ExcelFile('/home/ec2-user/book.xlsx')

#set dataframes for each sheet
df1 = pd.read_excel(data, '1-system.cpu.utilization', parse_cols = 'G,H,I')
df2 = pd.read_excel(data, '2-system.memory.usage.physical.', parse_cols = 'G,H,I')
df3 = pd.read_excel(data, '3-system.disk.used.util.percent', parse_cols = 'G,H,I')
df4 = pd.read_excel(data, '1-system.cpu.utilization', parse_cols = 'A,B')

df1 = df1[df1['Minimum'].notnull()]
df2 = df2[df2['Minimum'].notnull()]
df3 = df3[df3['Minimum'].notnull()]
df4 = df4[df4['Resource Name'].notnull()]

#combine sheets with formating
con = pd.concat([df4,df1.reset_index(drop=True),df2.reset_index(drop=True),df3.reset_index(drop=True)],axis = 1, keys = ['','CPU','Memory','Disk'],join='inner')

ht = con.to_html('ht.html')

#use this line only for Windows OS: 
config = pdf.configuration(wkhtmltopdf="/usr/local/bin/wkhtmltopdf")
#.

option ={'page-size':'A4', 'orientation':'Landscape'}

#convert to pdf
pdf.from_file('ht.html', 'output.pdf', configuration = config, options = option)
