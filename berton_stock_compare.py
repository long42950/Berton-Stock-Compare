import pandas as pd
import numpy as np
import time
import sys
import random
from enum import Enum
from termcolor import colored

class Version(Enum):
    OLD = "Old"
    NEW = "New"
    
berton_old = pd.DataFrame()
berton_new = pd.DataFrame()
code = []
des = []
hand = []
restock = []
version = []
get_old_stock = False

def get_excels():
    global berton_old,berton_new,get_old_stock
    need_file = True
    while need_file:
        oldName = input("Enter name of old Berton stock sheet: ")
        name = "./" + oldName + ".xlsx"
        try:
            berton_old = pd.read_excel(name)
            need_file = False
        except FileNotFoundError as error:
            print(f"File not found. Make sure the name is correct and the file and placed at the same folder as the script!")
    need_file = True
    while need_file:
        newName = input("Enter name of new Berton stock sheet: ")
        name = "./" + newName + ".xlsx"
        try:
            berton_new = pd.read_excel(name)
            need_file = False
        except FileNotFoundError as error:
            print("File not found. Make sure the name is correct and the file and placed at the same folder as the script!")
    need_file = True
    while need_file:
        response = input("Get listing from old list with restock ETA?[y/n]: ")
        if response == "y":
            get_old_stock = True
            need_file = False
        elif response == "n":
            need_file = False

def add_to_list(index, from_old):
    global berton_new,berton_old,code,des,hand,restock, version

    current_code = berton_old["Code"][index] if from_old else berton_new["Code"][index] 
    current_hand = berton_old["On Hand"][index] if from_old else berton_new["On Hand"][index]
    current_restock = berton_old["Estimated Restocking Date"][index] if from_old else berton_new["Estimated Restocking Date"][index]
    current_des = berton_old["Description"][index] if from_old else berton_new["Description"][index]
    check_progress(index)
    code.append(current_code)
    des.append(current_des)
    hand.append(current_hand)
    restock.append(current_restock)
    version.append(Version.OLD.value if from_old else Version.NEW.value)
    
def find_listing():
    global berton_new,berton_old,code,des,hand,restock
    i = len(berton_new["Code"])
    j = len(berton_old["Code"])
    for index in range(i):
        found = False
        current_code = berton_new["Code"][index]
        current_hand = berton_new["On Hand"][index]
        current_restock = berton_new["Estimated Restocking Date"][index]
        if index <= i//2:
            for jndex in range(j):
                check_code = berton_old["Code"][jndex]
                check_hand = berton_old["On Hand"][jndex]
                check_restock = berton_old["Estimated Restocking Date"][jndex]
                if check_code == current_code:
                    if check_hand != current_hand or type(current_restock) != float or type(check_restock) != float:
                        if get_old_stock:
                            add_to_list(jndex, True)
                        break     
                    found = True
                    break
        else:
            for jndex in reversed(range(j)):
                check_code = berton_old["Code"][jndex]
                check_hand = berton_old["On Hand"][jndex]
                check_restock = berton_old["Estimated Restocking Date"][jndex]
                if check_code == current_code:
                    if check_hand != current_hand or type(current_restock) != float or type(check_restock) != float:
                        if get_old_stock:
                            add_to_list(jndex, True)
                        break     
                    found = True
                    break
        if not found:
            add_to_list(index,False)
           

def create_excel():
    global berton
    
    new = pd.DataFrame({"Code": code, "Description": des, "On Hand": hand, "restock ETA": restock, "Version": version})
    while True:
        try:
            new.to_excel("./output.xlsx", sheet_name="Berton Stocks")
            print("New excel generated.")
            return
        except PermissionError as error:
            print(error)
            print("Make sure you don't have the output file opened.")
            input("Press any key to retry...")

def edit_excel():
    try:
        writer = pd.ExcelWriter("./ouput.xlsx", engine='xlsxwriter')
    except FileNotFoundError as error:
        print(error)
        return

def replace():
    size = 10
    true_output = ""
    for j in range(10000):         
        output = true_output
        for i in range(size):
            output = output + chr(random.randint(33,100))
            sys.stdout.write(output + "\r" if j != 10000 else "")
            if j != 10000:
                sys.stdout.flush()
        if j % 1000 == 0:
            true_output = true_output + chr(random.randint(65,90))
            size = size - 1
    sys.stdout.write(true_output+"\n")

def check_progress(index):
    size = len(berton_new["Code"])
    output = str(index)+" out of " + str(size) + " product checked..."
    sys.stdout.write(output+"\r" if index != size else "")
    sys.stdout.flush()
    

def percentage():
    size = 250000
    for i in range(size):
        progress = i / size * 100
        output = str(i)+" out of " + str(size) + " product checked..."
        sys.stdout.write(output+"\r" if i != size else "")
        if i != size:
                sys.stdout.flush()         
    print("All products checked...")   

def select_service():
    while True:
        print("                 ===========================================")
        print("                 |       Berton Stock Update Service       |")
        print("                 |       1. Get combined list              |")
        print("-----------------|       2. Apply Spreadsheet style        |-----------------")
        print("                 |       3. Replace                        |")
        print("                 |       4. Percentage                     |")
        print("                 ===========================================\n")
        service = input("Select the service you need: ")
        match service:
            case "1":
                get_excels()
                find_listing()
                print("All product has been checked.    ")
                create_excel()
                input("Press any Key to continue.")     
            case "2":
                print("Service not available!")                
            case "3":
                replace()                
            case "4":
                percentage()                
            case _:
                print(colored("Invalid Input!!!\n", 'red'))
            

select_service()
#berton = pd.read_excel("./tips.xlsx", index_col=0)
#berton = pd.read_excel("./berton05112021.xlsx")
#print(len(berton["Code"]))
#create_excel()


