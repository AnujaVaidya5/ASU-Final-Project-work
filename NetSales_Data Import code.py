#!/usr/bin/env python
# coding: utf-8

# In[29]:


import xlrd
import pymysql

Sales_data = xlrd.open_workbook("C:\Dailysalesforanalysis_2.xls")
sheet = Sales_data.sheet_by_index(0)

database = pymysql.connect(host="localhost", user="root", passwd= "Anuja@123", db="sales_data")
cursor = database.cursor()
query = """INSERT INTO Trans_NetSales(Sales_Day, Sales_Date, Weather, Customer_Count,Average_Check_value, Net_Sales, DrvThr, DrvThr_percent, Employee_Discount,Senior_Discount, Refunds, Mgr_discount, Labor_hrs) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
                           
for r in range(3, sheet.nrows):
        Sales_Day = sheet.cell(r,0).value
        Sales_Date = sheet.cell(r,1).value
        Weather = sheet.cell(r,2).value
        Customer_Count = sheet.cell(r,3).value
        Average_Check_value = sheet.cell(r,4).value
        Net_Sales = sheet.cell(r,5).value
        DrvThr = sheet.cell(r,6).value
        DrvThr_percent = sheet.cell(r,7).value
        Employee_Discount = sheet.cell(r,8).value
        Senior_Discount = sheet.cell(r,9).value
        Refunds = sheet.cell(r,10).value
        Mgr_discount = sheet.cell(r,11).value
        Labor_hrs = sheet.cell(r,12).value
                           
        values  = (Sales_Day, Sales_Date, Weather, Customer_Count,Average_Check_value, Net_Sales,DrvThr, DrvThr_percent, Employee_Discount,Senior_Discount, Refunds, Mgr_discount, Labor_hrs)                   
                           
        cursor.execute(query,values)    
        
cursor.close()
database.commit()
database.close()      

print(str(sheet.ncols))
print(str(sheet.nrows))
                           


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




