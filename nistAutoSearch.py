# -*- coding: utf-8 -*-
"""
Created on Mon Jun  3 11:21:29 2024

@author: Alonso
"""

import os 
import subprocess
import time
import numpy as np
import pandas as pd 

FULL_PATH_TO_MAIN_LIBRARY = "C:\\NISTDEMO\\MSSEARCH"
FULL_PATH_TO_WORK_DIR = "C:\\MS\\DATA\\RDA_Target_NIST.txt" #THE FOLDER WHERE THE RDA IS 


def removeLast(string):
    
    words = string.split('\\')
    words.pop()
    new = '\\'.join(words)
    
    return new


FULL_PATH_NO_TXT = removeLast(FULL_PATH_TO_WORK_DIR)

with open (FULL_PATH_TO_MAIN_LIBRARY +'\\AUTOIMP.MSD', 'w') as file:
    
    file.write(FULL_PATH_NO_TXT+'\\second.txt')  #try this except ERROR NOT FILE FOUND, \\ ADDED? MAKE SURE RIGHT ADDRESS

#PARA VERIFICAR QUE LO QUE ESCRIBES ESTA BIEN 
# with open (FULL_PATH_TO_MAIN_LIBRARY +'\\AUTOIMP.MSD', 'r') as file:
    
#     auto_file = file.readline().strip()

import ctypes

user32 = ctypes.windll.user32
hwnd = user32.GetForegroundWindow()
print(f"Handle of the current active window: {hwnd}")



with open(FULL_PATH_NO_TXT+'\\second.txt', 'w') as f:  #FULL_PATH_TO_WORK_DIR + 'second.txt'
    
    f.write(FULL_PATH_TO_WORK_DIR + f' APPEND\n{hwnd}')   #HAY QUE ESCRIBIR MAS AQUI FALTA EL OVERWRTIE EL 10 Y E L724 O ALGO DE ESO 
    


pathh = "C:\\NISTDEMO\\MSSEARCH\\nistms$.exe"
# print(out, err)
#command instrument activates the search window, write alone and then with PAR?
#FUNCIONAAA SI NO PONES EL SHEEEELL
# R = subprocess.run([f"{pathh}", '/instrument','/PAR=4'], capture_output= True, text = True)
R = subprocess.run([f"{pathh}", '/instrument','/PAR=2'], capture_output= True, text = True)

print(R.returncode, f'  error: {R.stderr}', R.stdout)


time.sleep(5)
c=1 
print('Start loading...')
while True:
    
#     result = subprocess.run(["cd", f"{FULL_PATH_TO_MAIN_LIBRARY}\\nistms$.exe", "/instrument", "/PAR=4"], shell = True, capture_output= True, text=True)
    if  os.path.exists(FULL_PATH_TO_MAIN_LIBRARY+'\\SRCREADY.TXT'):
        print ('Success!! Hits retrieved!')
        break
    else:
        time.sleep(10)
        print(f'Retrieving hits...{c}0 s')
        c+=1


# path = "C:\\NISTDEMO\\MSSEARCH\\"

# 
# # data = pd.read_csv(path + 'NISTLOG.TXT', header = 0, delimiter = ':', nrows=2)
# # data = pd.read_csv(path + 'NISTLOG.TXT', header = 0, usecols= ['Unknown'], nrows = 1)
# log_list = [['Mets',[]],['Label', []]]
# c=0 #counter
# with open(FULL_PATH_TO_MAIN_LIBRARY + 'NISTLOG.TXT' , 'r') as f:
    
#      for line in f:
#     #     # if number == 57:
#             # x = f.readline()
#              x = line.strip() #quita la \n del final 
#              if x.find('Unknown:') != -1:
                
#                  gamma1, *gamma2_l = x.strip('Unknown: ').split('             ')[0].split('   ')
#                  gamma2 = gamma2_l[0]
#                  mist = 'Compound' #possible string attached
#                  if gamma2.find(mist) != -1: #filter to remove possible string attached
#                      gamma2 = gamma2[:gamma2.find(mist)].strip()
             
#             # # dejar comentado
#             #     log_list[0][1].append(gamma1)
#             #     log_list[1][1].append(gamma2)
            
#              if x.find('Hit') != -1:
#                 hit_str = x.replace(':', ';', 1).replace(' ', ':', 1)
#                 beta = [string.split(':') for string in hit_str.split(';')]
                
#                 beta[1].insert(0, 'Name')
#                 beta[2].insert(0, 'Formula')
#                 beta.insert(0, ['Mets', gamma1]) #GAMMA 
#                 beta.insert(1, ['Label', gamma2]) #GAMMA 2
                
#                 # log_list[0][1].append(2)
#                 # log_list[1][1].append(3)
                
#                 if c == 0:
#                     beta = [[element[0],[element[1]]] for element in beta]
#                     log_list.extend(beta)
#                     log_dict = {element[0]:element[1] for element in log_list}  
                
#                 if c!= 0:
#                     for item in beta:
#                         key = item[0]
#                         value = item[1]
#                         if key in log_dict:
#                             log_dict[key].append(value)
#                 c = 1
#                 # log_list.extend()
                
#                 # beta[1] = [] ]
                
                

# # d = dict(log_list)
# df = pd.DataFrame(log_dict) #index = [0])

# df.to_excel('out_hits2.xlsx')

