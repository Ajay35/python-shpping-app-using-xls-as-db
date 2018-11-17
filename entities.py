import os
import sys
import collections
import xlrd
import xlwt
import matplotlib.pyplot as plt
from xlutils.copy import copy

class Customer:
    def __init__(self,cust_id,cust_name,cust_addr,phone_no):
        self.__id=cust_id
        self.__cust_name = cust_name
        self.__cust_addr=cust_addr
        self.__phone_np=phone_no

    def get_cust_id(self):
        return self.__id

    def set_customer_id(self,id):
        self.__id=id
        
    def get_customer_name(self):
        return self.__cust_name

    def set_customer_name(self,cname):
        self.__cust_name=cname
    
    def get_phone_no(self):
        return self.__phone_no

    def set_phone_no(self,pno):
        self.__phone_no=pno


    
class Product:
    def __init__(self,id,product_name,price,quantity,category):
        self.__id=id
        self.__product_name = product_name
        self.__price=price
        self.__quantity=quantity
        self.__category=category
        
    def get_product_id(self):
        return self.__id

    def set_product_id(self,id):
        self.__id=id
        
    def get_product_name(self):
        return self.__product_name

    def set_product_name(self,pname):
        self.__product_name=pname
    
    def get_price(self):
        return self.__price

    def set_price(self,p):
        self.__price=p

    def get_quantity(self):
        return self.__quantity

    def set_quantity(self,q):
        self.__quantity=q

    def get_category(self):
        return self.__category

    def set_quantity(self,cat):
        self.__category=cat
    
class Cart:
    items=[]
    def __init__(self,cart_id,no_of_items,total):
        self.__id=cart_id
        self.__no_of_items = no_of_items
        self.__total=total
        
    def set_id(self,id):
        self.__cart_id=id

    def get_id(self):
        return self.__cart_id
    
    def set_no_of_items(self,noofi):
        self.__no_of_items=noofi

    def get_no_of_items(self):
        return self.__no_of_items
    
    def set_total(self,t):
        self.__total=t

    def get_total(self):
        return self.__total

    def addItem(self,p):
        self.items.append(p)

## customer process and its helper functions
    
def customer_process(cust_name,cust_id):
    print "*******************"
    print "welcome to d-mart,"+cust_name
    list=[]
    new_cart=Cart(cust_id,0,0)
    while True:
        print "*************"
        print "1.View Products"
        print "2.Buy product"
        print "3.Add to cart"
        print "4.Delete from cart"
        print "5.View cart content"
        print "6.Checkout"
        print "7.View past orders"
        print "8.Exit to main menu"
        print "************"
        choice=input("Enter your choice:")
        if choice==1:
            view_products()

        elif choice==2:
            buy_product(cust_id,cust_name)
            
        elif choice==3:
            add_to_cart(new_cart)
            print "No of items in cart:",
            print new_cart.get_no_of_items()

        elif choice==4:
            if new_cart.get_no_of_items()==0:
                print "Cart is already empty"
            else:
                delete_from_cart(cust_name,new_cart)

        elif choice==5:
            view_cart_content(cust_name,new_cart)
            
        elif choice==6:
            if new_cart.get_no_of_items()==0:
                print "Cart is empty..."
                print "Press any key to continue"
                sys.stdin.read(1)
            else:
                checkout(cust_id,cust_name,new_cart)
                del new_cart.items[:]
                new_cart.set_no_of_items(0)

        elif choice==7:
            view_past_orders(cust_name,cust_id)
                
        elif choice==8:
            del new_cart
            break
        
        else:
            print "Invalid choice"

def buy_product(cust_id,cust_name):
    print "****************"
    c=input("Enter product id to buy:")
    workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
    worksheet = workbook.sheet_by_index(2)
    found=False
    found_at=0
    for row in range(worksheet.nrows):
        if worksheet.cell(row,0).value==c:
            if worksheet.cell(row,3).value > 0:
                found=True
                break
            else:
                print "Not in stock"
                sys.stdin.read(1)
                return None
        found_at+=1
        
    if found==True:
        id=worksheet.cell(found_at,0).value
        q=input("Enter quantity:")
        if q >= worksheet.cell(found_at,3).value:
            print "Not enough stock,Press any key to continue.."
            sys.stdin.read(1)
            return None
        card_no=input("Enter card no:")
        card_type=raw_input("Enter card type:")
        name=worksheet.cell(found_at,1).value
        price=worksheet.cell(found_at,2).value
        category=worksheet.cell(found_at,4).value
        prod=Product(id,name,price,q,category)
        wb=copy(workbook)
        sheet = wb.get_sheet(2)
        updated_q=worksheet.cell(found_at,3).value-q
        sheet.write(found_at,3,updated_q)
        wb.save('shopping_db.xls')
        
        print "*********************"
        print ""
        print "Item purchased:",
        print prod.get_product_name()
        print "Category",
        print prod.get_category()
        print "Quantity:",
        print prod.get_quantity()
        print "Each item price:",
        print prod.get_price()
        print "Total:",
        print prod.get_quantity()*prod.get_price()
        print "\n"
        print "Thank you for shopping with us,",
        print cust_name
        print "*********************"

        workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
        worksheet = workbook.sheet_by_index(0)
        found=False
        found_at=0
        for row in range(worksheet.nrows):
            if worksheet.cell(row,0).value==cust_id:
                    found=True
                    break
            found_at+=1
        wb=copy(workbook)
        sheet = wb.get_sheet(0)
        updated_exp=worksheet.cell(found_at,5).value+(prod.get_price()*prod.get_quantity())
        sheet.write(found_at,5,updated_exp)
        wb.save('shopping_db.xls')


        workbook1= xlrd.open_workbook('shopping_db.xls',on_demand=True)
        worksheet1 = workbook1.sheet_by_index(3)
        wb=copy(workbook1)
        status="pending"
        sheet = wb.get_sheet(3)
        sheet.write(worksheet1.nrows,0,cust_id)
        sheet.write(worksheet1.nrows,1,cust_name)
        sheet.write(worksheet1.nrows,2,prod.get_quantity()*prod.get_price())
        sheet.write(worksheet1.nrows,3,card_no)
        sheet.write(worksheet1.nrows,4,card_type)
        sheet.write(worksheet1.nrows,5,status)
        wb.save('shopping_db.xls')
        sys.stdin.read(1)
    else:
        print "Invalid product ID"
        sys.stdin.read(1)


        
    
def add_to_cart(new_cart):
    c=input("Enter id of product to add in cart:")
    workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
    worksheet = workbook.sheet_by_index(2)
    found=False
    found_at=0 
    for row in range(worksheet.nrows):
        if worksheet.cell(row,0).value==c:
            if worksheet.cell(row,3).value > 0:
                found=True
                break
            else:
                print "Not enough stock"
                sys.stdin.read(1)
                return None
        found_at+=1
    if found==True:
        id=worksheet.cell(found_at,0).value
        name=worksheet.cell(found_at,1).value
        price=worksheet.cell(found_at,2).value
        category=worksheet.cell(found_at,4).value
        prod=Product(id,name,price,1,category)
        new_cart.addItem(prod)
        new_cart.set_no_of_items(new_cart.get_no_of_items()+1)
        print "added.."
        sys.stdin.read(1)
    else:
        print "check the product ID entered,Press any key to continue.."
        sys.stdin.read(1)

def delete_from_cart(cust_name,new_cart):
    flag=True
    while flag:
        view_cart_content(cust_name,new_cart)
        c=input("Enter product id to remove from cart")
        for a in new_cart.items:
            if a.get_product_id()==c:
                new_cart.items.remove(a)
                print "removed"
                break
        view_cart_content(cust_name,new_cart)
        cont=input("Press 0 to exit and 1 to continue")
        if cont==0:
            flag=False



            
def view_cart_content(cust_name,new_cart):
    print "***************"
    print "Dear ",
    print cust_name
    if new_cart.get_no_of_items()==0:
        print "cart empty,continue shopping.."
        print "Press any key to continue"
        sys.stdin.read(1)
        return None
    print "Current cart items:"
    print "|  ID  |",
    print "| Item name |",
    print "Price |",
    print "quantity |",
    print "category |"
    for a in new_cart.items:
        print " | ",
        print a.get_product_id(),
        print " | ",
        print a.get_product_name(),
        print " | ",
        print a.get_price(),
        print " | ",
        print a.get_quantity(),
        print " | ",
        print a.get_category(),
        print " | "
    print "********************"

def checkout(cust_id,cust_name,new_cart):
    total=0
    list_items=[]
    card_no=input("Enter card no:")
    card_type=raw_input("Enter card type:")
    print "Current cart items:"
    print "| ID ",
    print "| Item name |",
    print "Price |",
    print "quantity |",
    print "category |"
    for a in new_cart.items:
        total+=a.get_price()
        print "| ",
        print a.get_product_id(),
        print "| ",
        print a.get_product_name(),
        print "| ",
        print a.get_price(),
        print "| ",
        print a.get_quantity(),
        print "| ",
        print a.get_category(),
        print "| "
        list_items.append(a.get_product_id())
    print "********************"
    print "Total:",
    print total

    print "Thank you for shopping with us...."


    workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
    worksheet = workbook.sheet_by_index(2)
    found_at=[]
    for id in list_items:
        found=0
        for row in range(worksheet.nrows):
            if worksheet.cell(row,0).value==id:
                found_at.append(found)
                break
            found+=1
    unique_occ=collections.Counter(found_at)
    found_at=list(set(found_at))
    wb=copy(workbook)
    sheet = wb.get_sheet(2)

    for l in found_at:   
        sheet.write(l,3,worksheet.cell(l,3).value-unique_occ[l])
    wb.save('shopping_db.xls')

    workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
    worksheet = workbook.sheet_by_index(0)
    f=0
    for row in range(worksheet.nrows):
        if worksheet.cell(row,0).value==cust_id:
            wb=copy(workbook)
            sheet = wb.get_sheet(0)
            sheet.write(f,5,worksheet.cell(f,5).value+total)
            break
        f+=1
    wb.save('shopping_db.xls')

    workbook1 = xlrd.open_workbook('shopping_db.xls',on_demand=True)
    worksheet1 = workbook.sheet_by_index(3)
    wb=copy(workbook)
    status="pending"
    sheet = wb.get_sheet(3)
    sheet.write(worksheet1.nrows,0,cust_id)
    sheet.write(worksheet1.nrows,1,cust_name)
    sheet.write(worksheet1.nrows,2,total)
    sheet.write(worksheet1.nrows,3,card_no)
    sheet.write(worksheet1.nrows,4,card_type)
    sheet.write(worksheet1.nrows,5,status)
    wb.save('shopping_db.xls')

def view_past_orders(cust_id,cust_name):
    workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
    worksheet = workbook.sheet_by_index(3)
    f=0
    for row in range(worksheet.nrows):
        if worksheet.cell(row,0).value==cust_id:
            wb=copy(workbook)
            sheet = wb.get_sheet(0)
            sheet.write(f,5,worksheet.cell(f,5).value+total)
            break
        f+=1
    wb.save('shopping_db.xls')








    
## admin processs and its helper functions   
    
def admin_process(admin_name):
    print "***********************"
    print "welcome,"+admin_name
    while True:
        print "*************"
        print "1.View products"
        print "2.Add product"
        print "3.Delete product"
        print "4.Modify Product"
        print "5.View customer"
        print "6.Graph of customer purchase histors"
        print "7.Exit to main menu"
        print "*************"
        op=input("select operation:")
        #os.system('cls' if os.name == 'nt' else 'clear')
        if op==1:
            view_products()
        elif op==2:
            add_product()
        elif op==3:
            delete_product()
        elif op==4:
            modify_product()
        elif op==5:
            view_customers()
        elif op==6:
            plot_graph()
        elif op==7:
            break
        else:
            print "Invalid operations"

def view_products():
     workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
     worksheet = workbook.sheet_by_index(2)
    
     for row in range(worksheet.nrows):
         for col in range (worksheet.ncols):
             print "|",
             print worksheet.cell(row,col).value,
             print "|",
         print "\n"
         
             
def add_product():
     workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
     worksheet = workbook.sheet_by_index(2)
     list=[]
     for r in range(worksheet.nrows):
         if worksheet.cell(r,0).value=="id":
             continue
         print type(worksheet.cell(r,0).value)
         list.append(worksheet.cell(r,0).value)

     id=max(list)+1
     product_name=raw_input("Enter product name:")
     price=input("Enter price:")
     quantity=input("Enter quantity:")
     category=raw_input("Enter category:")
     wb=copy(workbook)
     sheet = wb.get_sheet(2) 
     sheet.write(worksheet.nrows,0,id)
     sheet.write(worksheet.nrows,1,product_name)
     sheet.write(worksheet.nrows,2,price)
     sheet.write(worksheet.nrows,3,quantity)
     sheet.write(worksheet.nrows,4,category)
     wb.save('shopping_db.xls')
     print 'added product successfully'


def delete_product():
    id=input("Enter the id of product to delete:")
    workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
    worksheet = workbook.sheet_by_index(2)
    save_list=[]
    for row in range(worksheet.nrows):
        if worksheet.cell(row,0).value==id:
            continue
        else:
            print worksheet.cell(row,0).value
            prod_id=worksheet.cell(row,0).value
            name=worksheet.cell(row,1).value
            price=worksheet.cell(row,2).value
            quantity=worksheet.cell(row,3).value
            category=worksheet.cell(row,4).value
            x=Product(prod_id,name,price,quantity,category)
            save_list.append(x)
        
    wb=copy(workbook)
    sheet = wb.get_sheet(2)

    for r in range(worksheet.nrows):
        for c in range(worksheet.ncols):
            sheet.write(r,c,None)
    r=0
    for obj in save_list:
        sheet.write(r,0,obj.get_product_id())
        sheet.write(r,1,obj.get_product_name())
        sheet.write(r,2,obj.get_price())
        sheet.write(r,3,obj.get_quantity())
        sheet.write(r,4,obj.get_category())
        r+=1
    wb.save('shopping_db.xls')
    print 'deleted product successfully'

    
def modify_product():
    id=input("Enter id of product to modify:")
    workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
    worksheet = workbook.sheet_by_index(2)
    found=False
    found_row=0
    for row in range(worksheet.nrows):
        if worksheet.cell(row,0).value==id:
            found=True
            break
        else:
            found_row+=1
    if found==True:
        prod_name=raw_input("Enter new product name:")
        price=input("Enter new price:")
        quantity=input("Enter new quantity:")
        category=raw_input("Enter new category:")
        wb=copy(workbook)
        sheet = wb.get_sheet(2)
        sheet.write(found_row,1,prod_name)
        sheet.write(found_row,2,price)
        sheet.write(found_row,3,quantity)
        sheet.write(found_row,4,category)
        print "Product updated successfully"
        wb.save('shopping_db.xls')
    else:
        print "Product not found"

def plot_graph():
     workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
     worksheet = workbook.sheet_by_index(0)
     x_co=[]
     y_co=[]
     r=0
     for row in range(worksheet.nrows):
         if r==0:
             r+=1
             continue
         x_co.append(worksheet.cell(row,0).value)
         y_co.append(worksheet.cell(row,5).value)

     fig=plt.figure()    
     plt.plot(x_co, y_co, 'ro')
     fig.suptitle('Customer Purchase History', fontsize=20)
     plt.xlabel('Cutomer IDs', fontsize=18)
     plt.ylabel('Purchase(In Rupees)', fontsize=16)
     fig.savefig('test.png')
     plt.show()
   
    
def view_customers():
    workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
    worksheet = workbook.sheet_by_index(0)
    print "| ID ",
    print "| Name",
    print "| Address",
    print "| Phone No.",
    print "| Expenses",
    print " |"
    for row in range(worksheet.nrows):
        if row==0:
            continue
        print " |  ",
        print worksheet.cell(row,0).value,
        print " |  ",
        print worksheet.cell(row,1).value,
        print " |  ",
        print worksheet.cell(row,3).value,
        print " |  ",        
        print worksheet.cell(row,4).value,
        print " |  ",
        print worksheet.cell(row,5).value,
        print " |  "

## Guest process and helper functions

def guest_process():
    print "*************"
    print "Welcome guest"
    while True:
        print "*********"
        print "1.View products"
        print "2.Register"
        print "3.Exit to main menu"
        c=input("Enter choice:")
        if c==1:
            view_products()
            
        elif c==2:
            register()
            break
            
        elif c==3:
            break
            
        else:
            print "Unrecognised operation"


def register():
     workbook = xlrd.open_workbook('shopping_db.xls',on_demand=True)
     worksheet = workbook.sheet_by_index(0)
     list=[]
     for r in range(worksheet.nrows):
         if worksheet.cell(r,0).value=="cust_id":
             continue
         list.append(worksheet.cell(r,0).value)

     id=max(list)+1
     name=raw_input("name:")
     passwd=input("password (digits only):")
     address=raw_input("address:")
     phone=input("phone no.:")
     expense=0
     wb=copy(workbook)
     sheet = wb.get_sheet(0) 
     sheet.write(worksheet.nrows,0,id)
     sheet.write(worksheet.nrows,1,name)
     sheet.write(worksheet.nrows,2,passwd)
     sheet.write(worksheet.nrows,3,address)
     sheet.write(worksheet.nrows,4,phone)
     sheet.write(worksheet.nrows,5,expense)
     wb.save('shopping_db.xls')
     print "Guest registered successfully with ID",
     print id
     print "Login using ID and password"
     sys.stdin.read(1)
