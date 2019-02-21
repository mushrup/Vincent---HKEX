#!/usr/bin/env python
# coding: utf-8

# In[2]:


#GUI
import PySimpleGUI as sg
def GUI_input(index):
    global file_name,sheet_name,date2add,lower_bound,upper_bound
    values = ["a","b","c","d","e"]
    layout = [[sg.Text('Excel File(eg:feb19):')],      
              [sg.Input()],
              [sg.Text('Sheet Name(eg:Feb 20190129-30):')],      
              [sg.Input()],
              [sg.Text('Date to Add(eg:190208):')],      
              [sg.Input()],
              [sg.Text('Lower Bound(eg:22000):')],      
              [sg.Input()],
              [sg.Text('Upper Bound(eg:30800):')],
              [sg.Input()],
              [sg.RButton('Read'), sg.Exit()]]      

    window = sg.Window('Dates and Bounds for '+index).Layout(layout)      
    while True:
        event, user_input = window.Read()      
        if event is None or event == 'Exit':      
            window.Close()
        else:
            values = user_input
            break
    window.Close()
    
    file_name = values[0]
    sheet_name = values[1]
    date2add = values[2]
    lower_bound = int(values[3])
    upper_bound = int(values[4])


# In[3]:


#read file from last trading day and update columns
import pandas as pd
def Update_Last_Day(file_name,sheet_name):
    global pd1
    pd1=pd.read_excel(file_name+'.xlsx',sheet_name,header=6,index_col=None)
    pd1['C'] = pd1['D']
    pd1['F'] = pd1['G']
    pd1['X'] = pd1['I']
    pd1['P'] = pd1['O']
    pd1['S'] = pd1['R']
    pd1['AF'] = pd1['M']
    pd1['D'] = pd1['A']
    pd1['R'] = pd1['U']
    pd1.set_index('K',inplace=True)


# In[4]:


#download file from latest trading day
from urllib import request
from lxml import etree
def URL_Extract(url_front,date2add):
    url=url_front+date2add+".htm"
    headers={}
    headers['User-Agent']="Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36"
    req = request.Request(url, headers = headers)
    req = request.urlopen(req).read().decode("big5",errors="ignore")
    sel=etree.HTML(req)
    web_text = sel.xpath(r"//body/pre/text()")[1]
    with open('html.csv','w',encoding='big5') as outfile:
        outfile.write(web_text)
        outfile.close()


# In[5]:


#update new columns in earlier file
def Update_to_Merge(pd1,gap,index):
    global pd5
    pd2 = pd.read_csv('html.csv',encoding='big5',delimiter='\s+',index_col=None)
    pd2.set_index("合約月份",inplace=True)
    pd3 = pd2[pd2['行使價']=='認購'].loc[str(lower_bound):str(upper_bound)]
    #C/P-行使價;price-合約月份;OI-未平倉合約
    #0-認購;1-認沽
    pd4 = pd2[pd2['行使價']=='認沽'].loc[str(lower_bound):str(upper_bound)]
    number_of_entries = int((upper_bound-lower_bound)/gap+1)
    for i in range(number_of_entries):
        pd1.at[lower_bound+i*gap,'A']= pd3['成交量.2'][i]
        pd1.at[lower_bound+i*gap,'G']= pd3['*合約最低'][i]
        pd1.at[lower_bound+i*gap,'I']= pd3['*全日最低.1'][i]
        pd1.at[lower_bound+i*gap,'U']= pd4['成交量.2'][i]
        pd1.at[lower_bound+i*gap,'O']= pd4['*合約最低'][i]
        pd1.at[lower_bound+i*gap,'M']= pd4['*全日最低.1'][i]

    #write to the finalized output file
    pd5 = pd.read_excel(file_name+'.xlsx',sheet_name)
    for i in range(number_of_entries):
        for j in range(34):
            if j == 10:
                continue
            if j < 10:
                pd5.iloc[6+i,[j]] = pd1.iloc[i,j]
            if j > 10:
                pd5.iloc[6+i,[j+1]] = pd1.iloc[i,j]
    pd5.to_csv(index+'_output.csv',sep=',',encoding='utf-8',index=False)


# In[6]:


#GUI2
def GUI_index(index):
    layout=[[sg.Text('Please check file:'+index+'_output.csv')],
           [sg.Text('Step 1: Open '+index+'_output.csv with Notepad')],
            [sg.Text('Step 2: \'File\' -> \'Save as\' under the same name (replace the original file)')],
            [sg.Text('Step 3: Right click '+index+'_output.csv and choose open with Excel')]]
    window = sg.Window('Read Me').Layout(layout)
    window.Show()


# In[ ]:


#HSI
GUI_input('HSI')
Update_Last_Day(file_name,sheet_name)
url_front="https://www.hkex.com.hk/chi/stat/dmstat/dayrpt/hsioc"
URL_Extract(url_front,date2add)
Update_to_Merge(pd1,200,"HSI")
GUI_index('HSI')


# In[ ]:


#HHI
GUI_input('HHI')
Update_Last_Day(file_name,sheet_name)#sheet_name:HHIO Feb 20190129-30
url_front="https://www.hkex.com.hk/chi/stat/dmstat/dayrpt/hhioc"
URL_Extract(url_front,date2add)
Update_to_Merge(pd1,100,"HHI")
GUI_index('HHI')


# In[ ]:




