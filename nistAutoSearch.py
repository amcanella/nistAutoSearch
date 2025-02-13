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
import ctypes
import datetime

#PATHS
FULL_PATH_TO_MAIN_LIBRARY = "C:\\NISTDEMO\\MSSEARCH" #THE FOLDER WHERE THE PROGRAM IS, PROGRAM DIRECTORY 
FULL_PATH_TO_WORK_DIR = "C:\\MS\\DATA\\RDA_Target_NIST.txt" #THE FOLDER WHERE THE RDA IS, WORK DIRECTORY

#INIT SETTINGS, IMPORTANT TO CLOSE NIST MS SEARCH PROGRAM EVERY TIME YOU MAKE A CHANGE 
number_hits = 5  #WRITE THE NUMBER OF HITS TO LOG
automatic = 1   #ACTIVATE THE AUTOMATION 

def flagChecker():
    #Opens INI file in the PROGRAM DIRECTORY and sets flags for automation and number of hits
    
    with open(f'{FULL_PATH_TO_MAIN_LIBRARY}\\nistms.INI', 'r') as ini:
        
        lines = ini.readlines()
        
    with open(f'{FULL_PATH_TO_MAIN_LIBRARY}\\nistms.INI', 'w') as ini:
        
        for line  in lines:
            if line.startswith('Hits to Print'):
                line = 'Hits to Print=' + str(number_hits)+'\n'
                # print(line)
                
            if line.startswith('Automatic'):
                line = 'Automatic=' + str(automatic)+'\n'
                # print(line)
            ini.write(line)


def removeLast(string):
    #Removes the txt file name from path
    
    words = string.split('\\')
    words.pop()
    new = '\\'.join(words)
    
    return new


def wndN():
    #Returns program main window handle numbers
    
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    #print(f"Handle of the current active window: {hwnd}")
    return hwnd


def createFiles():
    #Creates the locator files required to use the automation features
    
    with open (FULL_PATH_TO_MAIN_LIBRARY +'\\AUTOIMP.MSD', 'w') as file:
        
        file.write(FULL_PATH_NO_TXT+'\\second.txt')  
    
    with open(FULL_PATH_NO_TXT+'\\second.txt', 'w') as f:
        
        f.write(FULL_PATH_TO_WORK_DIR + f' OVERWRITE\n{hwnd}')


def exCmd():
    #Executes the PAR 2 commmand that carries out a background search and writes hit lists in log file
    
    R = subprocess.run([f"{FULL_PATH_TO_MAIN_LIBRARY}\\nistms$.exe", '/instrument','/PAR=2'], capture_output= True, text = True)
    
    print(f'Errors: {R.returncode}')
    
    time.sleep(5)
    c=1 
    print('Start loading...')
    while True:
        
        if  os.path.exists(FULL_PATH_TO_MAIN_LIBRARY+'\\SRCREADY.TXT'):
            print ('Success!! Hits retrieved!')
            break
        else:
            time.sleep(10)
            print(f'Retrieving hits...{c}0 s')
            c+=1


def lineChecker(line):
    #Checks whether the name within << >> has a semicolon
    
    checked = line
    start = line.find("<<")
    end = line.find(">>")
    colon = line.find(';', start, end)
    
    while colon != -1:
        
        checked = checked[:colon] + ' ' + checked[colon+1:]
        colon = checked.find(';', start, end)
        line = checked
        # print(colon)
        # print(checked)

    
        
    return checked


def logReader():
    #Reads the log files and returns hit dictionary
    
    print('Reading log files...')
    log_list = [['Mets',[]],['Label', []]]
    c=0 #counter
    with open(FULL_PATH_TO_MAIN_LIBRARY + '\\SRCRESLT.TXT' , 'r') as f:
        
          for line in f:
              
                  x = line.strip() #removes final \n 
                  if x.find('Unknown:') != -1:
                    
                      gamma1, *gamma2_l = x.strip('Unknown: ').split('             ')[0].split('   ')
                      gamma2 = gamma2_l[0]
                      mist = 'Compound' #possible string attached
                      if gamma2.find(mist) != -1: #filter to remove possible string attached
                          gamma2 = gamma2[:gamma2.find(mist)].strip()
                
                  if x.find('Hit') != -1:
                    hit_str = x.replace(':', ';', 1).replace(' ', ':', 1)
                    
                    hit_str_checked = lineChecker(hit_str) #Checks that name has no ';' in it
                    hit_checked = hit_str_checked.split(';') 
                    beta = [string.split(':') for string in hit_checked]
                    
                    beta[1].insert(0, 'Name')
                    beta[2].insert(0, 'Formula')
                    beta.insert(0, ['Mets', gamma1])  
                    beta.insert(1, ['Label', gamma2]) 
                    
                    if c == 0:
                        beta = [[element[0],[element[1]]] for element in beta]
                        log_list.extend(beta)
                        log_dict = {element[0]:element[1] for element in log_list}  
                    
                    if c!= 0:
                        for item in beta:
                            key = item[0]
                            if not item[1]:
                                raise ValueError("The second element is empty")
                            else:
                                value = item[1]
                            if key in log_dict:
                                log_dict[key].append(value)
                    c = 1
                    
    return log_dict


if __name__ ==  "__main__" :
    
    flagChecker()
    
    FULL_PATH_NO_TXT = removeLast(FULL_PATH_TO_WORK_DIR)
    now = datetime.datetime.now()
    nowf = now.strftime("%Y-%m-%d-%H-%M-%S")
    
    hwnd = wndN() #window handler number 
    
    createFiles()

    exCmd()
    
    d = logReader()
    
    df = pd.DataFrame(d)
    
    df.to_excel(f'{FULL_PATH_NO_TXT}\\out_hits_{nowf}.xlsx')
    print('Done! Your excel file is ready!')
    print(f'Hits stored in: {FULL_PATH_NO_TXT}\\out_hits_{nowf}.xlsx')
