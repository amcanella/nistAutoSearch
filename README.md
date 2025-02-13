# nistAutoSearch

Automated tool for NIST MS search, data retrieval and storage in excel file.  

### Instructions of use:

1. Download the latest version of NIST MS Search Program. You can install from: https://chemdata.nist.gov/dokuwiki/doku.php?id=chemdata:nistlibs

2. Update the path variables in the script with your own: 

```python
FULL_PATH_TO_MAIN_LIBRARY = "C:\\NISTDEMO\\MSSEARCH" #THE FOLDER WHERE THE PROGRAM IS, PROGRAM DIRECTORY 
FULL_PATH_TO_WORK_DIR = "C:\\MS\\DATA\\RDA_Target_NIST.txt" # LOCATION OF RDA FILE, WORK DIRECTORY
```

3. Update the number of hits to log (default 5) and make sure the automation is enabled(1):

```python
number_hits = 5  #WRITE THE NUMBER OF HITS TO LOG
automatic = 1   #ACTIVATE THE AUTOMATION 
```

4. Run the script. The hits are stored in an excel file in the 'FULL_PATH_TO_WORK_DIR' under 'out_hits_[date].xlsx' by default.


### Additional

- Only in case of error. This should be done automatically when downloading the program, however if it is not, the MS Search user guide suggests that: 

" It may be necessary to determine the location of the NIST MS Search Program files and the location of the NIST MS Search working directory. These may be found in the [NISTMS] section of the WIN.INI file located in the %windir% directory. Below is a typical example of this section, you must write your own: 

[NISTMS] 
Path32=C:\NIST17\MSSEARCH\ 
WorkDir32=C:\NIST17\MSSEARCH\ 
Amdis32Path=C:\NIST17\AMDIS32\ 
AmdisMSPath=C:\NIST17\AMDIS32\ 

The string beyond “Path32=” is the path to MS Search Program’s folder (location of the executable files); the string beyond “WorkDir32=” is the path to the MS Search Program’s working directory (location of the history, Hits, etc.). Two other items refer to AMDIS and its connection to the MS Search Program. "