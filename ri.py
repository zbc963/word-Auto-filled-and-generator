# -*- coding: utf-8-*-  
import win32com  
from win32com.client import Dispatch, constants  
      
 
template_path = 'C:\Users\Jzzz\Desktop\\volunteer.doc'  
 
store_path = 'C:\Users\Jzzz\Desktop\ '  
   
NewStr = 'Pauline Shen'
NewNum = '16.83' 
      
 
w = win32com.client.Dispatch('Word.Application')  

w.Visible = True  
w.DisplayAlerts = True  

doc = w.Documents.Open( FileName = template_path )  

w.Selection.Find.ClearFormatting()  
w.Selection.Find.Replacement.ClearFormatting()  
      
    
lst = ['Siqing Wang',
'Leitian Wang',
'Lunyi Xiong',
'FIONA',
'Iris Zhong',
'Leizhu Chen',
'Rachel ',
'PEIYING YIN',
'Stephanie',
'Sylvia',
'Chunru Xue',
'Xinran Zhang',
'Ziqi Xu',
'Kei Kwan Yu',
'Shuya Shi',
'Yueying Jin',
'Tianai Guo',
'Alma',
'Frank',
'Denise',
'Chaoyi Dai',
'Mengyi Yuan',
'Yongyi Chen',
'Annie Yu',
'Meng Sun',
'Xinlei Su',




]  
num = ['16.86',
'16.87',
'16.88',
'16.89',
'16.90',
'16.91',
'16.92',
'16.93',
'16.94',
'16.95',
'16.96',
'16.97',
'16.98',
'16.99',
'16.100',
'16.101',
'16.102',
'16.103',
'16.104',
'16.105',
'16.106',
'16.107',
'16.108',
'16.109',
'16.110',
'16.111'



]
order=[
'1'
,'2'
,'3'
,'4'
,'5'
,'6'
,'7'
,'8'
,'9'
,'10'
,'11'
,'12'
,'13'
,'14'
,'15'
,'16'
,'17'
,'18'
,'19'
,'20'
,'21'
,'22'
,'23'
,'24'
,'25'
,'26'
# ,'27'
# ,'28'
# ,'29'
# ,'30'
# ,'31'
# ,'32'
# ,'33'
# ,'34'
# ,'35'
# ,'36'
# ,'37'
# ,'38'
# ,'39'
# ,'40'
# ,'41'
# ,'42'
# ,'43'
# ,'44'
# ,'45'
# ,'46'
# ,'47'
# ,'48'
# ,'49'
# ,'50'
# ,'51'
# ,'52'
# ,'53'
# ,'54'
# '55',
# '56',
# '57',
# '58',
# '59',
# '60',
# '61',
# '62',
# '63',
# '64',
# '65',
# '66',
# '67',
# '68',
# '69',
# '70',
# '71'
# '73',
# '74',
# '75',
# '76',
# '77'

]
 
g=0
for g in range(27):  
        OldStr, NewStr = NewStr, lst[g]  
        w.Selection.Find.Execute(OldStr, False, False, False, False, False, True, 1, True, NewStr, 49)
        OldNum, NewNum = NewNum, num[g] 
        w.Selection.Find.Execute(OldNum, False, False, False, False, False, True, 1, True, NewNum, 49)
        doc.SaveAs(store_path+order[g] +'.doc')

doc.Close()  
w.Documents.Close()  
w.Quit()  

