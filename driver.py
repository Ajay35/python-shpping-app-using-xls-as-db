import os
import xlrd
from entities import *

def main():
    while True:
        print "*******************"
        print "1.Admin"
        print "2.Cutomer (Registered customers only)"
        print "3.Guest"
        print "4.Exit"
        text=input("Login:")
        if text==1:
            found=False
            id=input("Enter admin id:")
            passwd=input("Enter admin password:")
            workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
            worksheet = workbook.sheet_by_index(1)

            admin_name=""
            for row in range(worksheet.nrows):
                if id==worksheet.cell(row,0).value:
                    if passwd==worksheet.cell(row,1).value:
                       found=True
                       admin_name=worksheet.cell(row,2).value
                       break
                
            if found==True:
                admin_process(admin_name)
            else:
                print "Invalid id or password"

                
        elif text==2:
            id=input("Enter customer id:")
            passwd=input("Enter cutomer password:")
            workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
            worksheet = workbook.sheet_by_index(0)
            found=False
            cust_name=""
            for row in range(worksheet.nrows):
                if id==worksheet.cell(row,0).value:
                    if passwd==worksheet.cell(row,2).value:
                        cust_name=worksheet.cell(row,1).value
                        found=True
                        break
            if found==True:
                customer_process(cust_name,id)
            else:
                print "invalid id or password"
        elif text==3:
            guest_process()
            
        elif text==4:
            sys.exit()
            
        else:
            print "Invalid choice"
        
                
                

if __name__=='__main__':
    main()
