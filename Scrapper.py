from tkinter import *
import time
from tkinter import messagebox
from bs4 import BeautifulSoup
import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from tkinter import ttk,filedialog
#import numpy
import os
import webbrowser

# getting search query and embedding it into 'get' request.
global mw
def mainWindow():
    global mw,searchText,txt,f1
    
    mw=Tk()
    txt=StringVar()
    mw.geometry('1200x800')
    f1=Frame(mw,width=100,height=20)
    f1.grid(row=0,column=0)
    searchText=Entry(f1,textvariable=txt,font=('Times New Roman',18),width=20)
    searchText.grid(row=0,column=0,sticky=W,padx=20)
    Button(f1,text="Search",command=checkvalid,width=12).grid(row=0,column=1,pady=10)
    
def excelToGui():
    global my_tree
    print("tree made")
def openxl():
    os.startfile("C://Users/KALP/Desktop/Project/Results.xlsx")
    #dfprint ("failed to connect")
   # storing the obtained html text into soup variable.
def checkvalid():
    query=str(txt.get())
    
    print(query)
    if(query==""):
        messagebox.showwarning("ERROR","Please Search Valid Product")
    else:
        searchButton()
def display():
    selected=my_tree.focus()
    values=my_tree.item(selected,'values')
    print(values[4])
    webbrowser.open(values[4])
def searchButton():
    global dict,names,prices,ratings,images,my_tree
    try:
           # l1=Label(mw,text="please wait")
           # l1.grid(row=0,column=2)#
           # waitmsg=messagebox.showinfo("STATUS","Please wait while we fetch result")
           #mw.after(0)
            
            f=0;a=0;
            # getting search query and embedding it into 'get' request.
            query=str(txt.get()).strip().replace(" ","+")
            
            print(type(query))
            #query = input("Enter the product name: ").strip().replace(" ","+")
            print(type(query))
           

            #messagebox.showinfo("STATUS","please wait")
            
            # putting the query at proper place
            html_text = requests.get("https://www.flipkart.com/search?q="+query+"&otracker=search&otracker1=search&marketplace=FLIPKART&as-show=on&as=off").text
            l1=Label(f1, text="Scrapping. Please wait...")
            l1.grid(column=0, row=2)
            progress = ttk.Progressbar(f1, orient = 'horizontal', mode = 'determinate', length = 240)
            progress.grid(row=1, column=0)
            soup = BeautifulSoup(html_text, 'lxml')
            progress['value']=10
            mw.update_idletasks()
            # scrapping all the neccessary information as lists from the first page of the website requested.
            names = soup.find_all('div', class_='_4rR01T')
            prices = soup.find_all('div', class_='_30jeq3 _1_WHN1')
            ratings = soup.find_all('div', class_='_3LWZlK')
            images = soup.find_all('img', class_='_396cs4 _3exPp9')
            links = soup.find_all('a', class_='_1fQZEK')
            progress['value'] = 20
            mw.update_idletasks()

                # creating a dictionary.
            dict = {'Product Name':[], 'Rating':[], 'Price(₹)':[],  'Webstore':[],'Link':[],'Image':[]}

                    # adding all the info into dictionary.
            for i in range(0,len(prices)):
                    dict['Product Name'].append(names[i].text)
                    dict['Price(₹)'].append(int(prices[i].text[1:].replace(",","")))
                    dict['Rating'].append(ratings[i].text+" *")
                    dict['Image'].append(images[i]["src"])
                    dict['Webstore'].append("Flipkart")
                    dict['Link'].append("https://www.flipkart.com"+links[i]["href"])
                    
                    f+=1
            progress['value'] = 30
            mw.update_idletasks()
            my_frame=Frame(mw)
            my_frame.grid(row=1,column=0,pady=50,padx=20)
          
            tree_scroll=Scrollbar(my_frame)
            style=ttk.Style()
            style.configure("Treeview",background="#EAEDF0",foreground="black",rowheight=30)
            my_tree=ttk.Treeview(my_frame ,yscrollcommand=tree_scroll.set,selectmode="extended",height=20)

            my_tree['columns']=("Product","Price","Ratings","Website","shop")
            my_tree.column("#0",width=0,minwidth=0)
            my_tree.column("Product",width=800,minwidth=120)
            my_tree.column("Ratings",width=80,minwidth=50)
            my_tree.column("Price",width=50,minwidth=50)
            my_tree.heading("Product",text="Product Name",anchor=W)
            my_tree.heading("Ratings",text="Ratings")
            my_tree.heading("Price",text="Price(₹)")
            my_tree.column("Website",width=80,minwidth=50)
            my_tree.heading("Website",text="Website",anchor=W)
            my_tree.column("shop",width=0,minwidth=0)
            my_tree.heading("shop",text="shop",anchor=W)
            
           
            
            for items in range(0,len(names)):
                my_tree.insert(parent='',index=False,values=(names[items].text,prices[items].text[1:],ratings[items].text[0:4],"Flipkart","https://www.flipkart.com"+links[i]["href"]))

            progress['value'] = 40
            mw.update_idletasks()
            option = webdriver.ChromeOptions()
            option.add_argument('headless')
            driver = webdriver.Chrome('./webdriver/chromedriver', options=option)
            progress['value'] = 50
            mw.update_idletasks()
            driver.get('https://www.amazon.in/')
            driver.find_element(By.ID, 'twotabsearchtextbox').send_keys(query)
            driver.find_element(By.ID, 'nav-search-submit-button').click()

            progress['value'] = 60
            mw.update_idletasks()
            time.sleep(3)
            html_text = driver.page_source

            soup = BeautifulSoup(html_text, 'lxml')
            progress['value'] = 80
            mw.update_idletasks()


            names = soup.find_all('span', class_='a-size-medium a-color-base a-text-normal')
            prices = soup.find_all('span', class_='a-price-whole')
            ratings = soup.find_all('span', class_='a-icon-alt')
            images = soup.find_all('img', class_='s-image')
            links = soup.find_all('a', class_='a-link-normal s-underline-text s-underline-link-text s-link-style a-text-normal')
            def min(a, b):
                if (a>b):
                    return b
                else:
                    return a
            for i in range(0,min(len(prices),len(names))):
                #print (i)
                dict['Product Name'].append(names[i].text)
                dict['Price(₹)'].append(int(prices[i].text.replace(",","")))
                dict['Rating'].append(ratings[i].text[0:3]+" *")
                dict['Image'].append(images[i]['src'])
                dict['Webstore'].append("Amazon")
                dict['Link'].append("https://www.amazon.in"+links[i]['href'])
            
                a+=1
           
            for items in range(0,len(names)):
                my_tree.insert(parent='',index=False,values=(names[items].text,prices[items].text,ratings[items].text[0:3   ],"Amazon","https://www.amazon.in"+links[i]['href']))
            
            tree_scroll=Scrollbar(my_frame)
            #tree_scroll.pack(side=RIGHT,fill=Y)
            tree_scroll.config(command=my_tree.yview)
            my_tree.pack(fill=BOTH,expand=TRUE)
            #messagebox.showinfo("STATUS","Here are Results")
            
            progress['value'] = 100
            mw.update_idletasks()
            progress.destroy()
            l1.destroy()
         
            
            if(a==0 & f==0):
                print("Both Empty")
                messagebox.showerror("ERROR","Invalid Search No Result")
            elif(a==0):
                print ("amazon empty")
                messagebox.showwarning("Warning","No results on Amazon")

            elif(f==0):
                print ("flipkart empty")
                messagebox.showwarning("Warning","No results on Flipkart")
            # exporting all the content of dictionary into an xlsx file
            
            df = pd.DataFrame(dict)
            df= df.sort_values(by=['Price(₹)'])
            index=False
            df.to_excel('./Results.xlsx', sheet_name="Results",index=False)
            b1=Button(f1,text="Open Excel",command=openxl,anchor=CENTER)
            b1.grid(row=0,column=2,padx=20)
            b2=Button(f1,text="SHOP ",command=display).grid(row=0,column=3)
            print("DONE!")
            l1=Label(mw,text="").grid(row=0,column=2)#waitmsg=messagebox.showinfo("STATUS","Thank You for Waiting")          
            if(not(a==0&f==0)):
                excelToGui()
            
    except requests.ConnectionError as e :
        print ("Can't connect to network")
        html_text="No Response"
        messagebox.showerror("ERROR","Can't Connect To Network")
    except IndexError :
        if(a==0 ):
            print("amazon empty")
        if(f==0 ):
            print("flipkart empty")
        messagebox.showerror("ERROR","Invalid Search")
        print ("No result found")
        print(f )
        print(a)



def file_open():
    global my_tree
    print("check3")
    filename="Results.xlsx"
    if filename :
        print ("file found")
        try:
            
            filename=r"{}".format(filename)
            df=pd.read_excel(filename)
        except ValueError:
            myLabel.config(text="File Not Found")
        except FileNotFoundError :
            my_label.config(text="file not")
    #clear_tree()
    my_tree["columns"]=list(df.columns)
    my_tree["show"]="headings"
    for column in my_tree["column"]:
        my_tree.heading(column,text=column)
    df_rows=df.to_numpy().tolist()
    for row in df_rows:
        my_tree.insert("","end",values=row)
    my_tree.grid(row=1,column=0)
    print(type(df_rows))
#def selectItem():
    select=my_tree.focus()
    det=my_tree.item(select)
    print(det)


mainWindow()
mw.mainloop()