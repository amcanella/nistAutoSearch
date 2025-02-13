# nistAutoSearch

Automated tool for NIST MS search, data retrieval and storage in excel file.  

Instructions of use:

Download the latest version of NIST MS Search Program.

Update the path variables in the script (first lines). 

If you encounter errors, the user guide suggests that: 

" It may be necessary to determine the location of the NIST MS Search Program files and the location of the NIST MS Search working directory. These may be found in the [NISTMS] section of the WIN.INI file located in the %windir% directory. Below is a typical example of this section: 

[NISTMS] 
Path32=C:\NIST17\MSSEARCH\ 
WorkDir32=C:\NIST17\MSSEARCH\ 
Amdis32Path=C:\NIST17\AMDIS32\ 
AmdisMSPath=C:\NIST17\AMDIS32\ 

The string beyond “Path32=” is the path to MS Search Program’s folder (location of the executable files); the string beyond “WorkDir32=” is the path to the MS Search Program’s working directory (location of the history, Hits, etc.). Two other items refer to AMDIS and its connection to the MS Search Program. "