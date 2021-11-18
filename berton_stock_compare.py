import pandas as pd
import numpy as np
from enum import Enum

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
    print(current_code)
    code.append(current_code)
    des.append(berton_new["Description"][index])
    hand.append(current_hand)
    restock.append(current_restock)
    version.append(Version.OLD.value if from_old else Version.NEW.value)
    

# TODO
#      - if old list has ETA, add both old and new
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
            print("New excel generated!!!")
            return
        except PermissionError as error:
            print(error)
            print("Make sure you don't have the output file opened.")
            input("Press any key to retry...")

get_excels()
find_listing()
create_excel()
#berton = pd.read_excel("./tips.xlsx", index_col=0)
#berton = pd.read_excel("./berton05112021.xlsx")
#print(len(berton["Code"]))
#create_excel()
input("Press any Key to continue.")


