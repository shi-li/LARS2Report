'''
Change value of Output_Case accordingly by examples below
Case 1
Value more than 100 needs to be reported as 99.9
95.4 needs to be reported as 95.0
95.5 needs to be reported as 95.5
95.9 needs to be reported as 95.5

Case 2
Value more than 100 needs to be reported as 99.9
95.4 needs to be reported as 95.0
95.5 needs to be reported as 95.0
95.9 needs to be reported as 95.0
'''

Output_Case = 1

import os
import win32com.client

#Read the path of LIS and .docx files from CURRENT folder.
def FindPath()
    legal_path = os.path.join(os.getcwd(),"legal.LIS")
    permit_path = os.path.join(os.getcwd(),"permit.LIS")
    
    for item in os.listdir('.'):
        if item.find(".docx") != -1:
            summary_path = os.path.join(os.getcwd(),item)






legal_file = open(legal_path, 'r')
legal_list = []
for line in legal_file:
    legal_list.append(line)

legal_values = []
for i in [0, 1, 2, 3, 4]:
    legal_values.append(float(legal_list[16+2*i][62:67]))
    legal_values.append(float(legal_list[16+2*(5+i)][62:67]))   
    legal_values.append(legal_list[16+2*i][35])
    
legal_file.close()


permit_file = open(permit_path, 'r')
permit_list = []
for line in permit_file:
    permit_list.append(line)

permit_values = []
for i in [0, 1, 2, 3, 4, 5, 6, 7]:
    permit_values.append(float(permit_list[16+2*i][62:67]))          
    permit_values.append(float(permit_list[16+2*(8+i)][62:67]))
    permit_values.append(permit_list[16+2*i][35])    

permit_file.close()

#Function to process both the controlling stress and the load capactiy in ton.
def rounddown(x):
    if x == 'M':
        return 'Moment'
    elif x == 'V':
        return 'Shear'
    elif x == 'S':
        return 'Serviceability'
    elif float(x)-int(x) >= 0.5:
        return float(int(x)+0.5)
    else:
        return str(int(x)) + '.0'


# Start the transfer of values to Word Document with
# win32com.client

w = win32com.client.Dispatch('Word.Application')
w.Visible = 0  
w.DisplayAlerts = 0
doc = w.Documents.Open(FileName = summary_path)

table_legal = doc.Tables(2)
table_permit = doc.Tables(3)


for i in xrange(5):
    table_legal.Cell(4+2*i,3).Range.Text = rounddown(legal_values[3*i])
for i in xrange(5):
    table_legal.Cell(4+2*i,4).Range.Text = rounddown(legal_values[1+3*i])
for i in xrange(5):
    table_legal.Cell(5+2*i,2).Range.Text = rounddown(legal_values[2+3*i])

for i in xrange(8):
    table_permit.Cell(4+2*i,3).Range.Text = rounddown(permit_values[3*i])
for i in xrange(8):
    table_permit.Cell(4+2*i,4).Range.Text = rounddown(permit_values[1+3*i])
for i in xrange(8):
    table_permit.Cell(5+2*i,2).Range.Text = rounddown(permit_values[2+3*i])


w.ActiveDocument.Close(SaveChanges=True)
w.Quit()
