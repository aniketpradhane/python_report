##################################################################
# Author : Aniket Pradhane                                       # 
# EMP ID : 609                                                   # 
# Purpose : XLS file Processor in Python                         #
# Date : 25 April 2019                                           # 
##################################################################

#Required libraries:  Pandas, pdfkit
#Install tool: wkhtmltopdf using 'sudo apt-get install wkhtmltopdf' on linux


import pandas as pd
import warnings
import pdfkit as pdf

warnings.filterwarnings("ignore", category=FutureWarning)

#read excel file
data = pd.ExcelFile('book.xlsx')

#set dataframes for each sheet
df1 = pd.read_excel(data, '1-system.cpu.utilization', parse_cols = 'J,K,L')
df2 = pd.read_excel(data, '2-system.memory.usage.physical.', parse_cols = 'J,K,L')
df3 = pd.read_excel(data, '3-system.disk.used.util.percent', parse_cols = 'J,K,L')
df5 = pd.read_excel(data, '4-network.out', parse_cols = 'J,K,L')
df6 = pd.read_excel(data, '5-network.in', parse_cols = 'J,K,L')
df4 = pd.read_excel(data, '1-system.cpu.utilization', parse_cols = 'D,E')

df1 = df1[df1['Minimum'].isnull()]
df2 = df2[df2['Minimum'].isnull()]
df3 = df3[df3['Minimum'].isnull()]
df5 = df5[df5['Minimum'].isnull()]
df6 = df6[df6['Minimum'].isnull()]
df4 = df4[df4['Resource Name'].isnull()]

#combine sheets with formating
con = pd.concat([df4,df1.reset_index(drop=True),df2.reset_index(drop=True),df3.reset_index(drop=True),df5.reset_index(drop=True),df6.reset_index(drop=True)],axis = 1, keys = ['','CPU','Memory','Disk','Network_OUT','Network_IN'],join='inner')

ht = con.to_html('ht.html')

#use this line only for Windows OS: 
config = pdf.configuration(wkhtmltopdf="/usr/local/bin/wkhtmltopdf")
#.

option ={'page-size':'A4', 'orientation':'Landscape'}

#convert to pdf
pdf.from_file('ht.html', 'output.pdf', configuration = config, options = option)
