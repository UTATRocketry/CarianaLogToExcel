# CarianaLogToExcel

[Avionics] Python Script to export Carina log files to excel for easier analysis.

# Requirements

Openpyxl is a requirement for this script in order to modify and write to excel documents. Install by typing the following in CMD and pressing enter:	

```
> pip install openpyxl
```

# Limitation and Usage

There are 3 modifiable parameters within the python script:

* Values Per Second (VPS)
* Excel File Name (excel_file_name)
* Names of all log files to be read by the program stored as list (logfile_names).

**Values Per Second:**

By default values per second appears as follows within the program:

```
VPS = 10
```

Meaning 1 values per sensor node are writen to excel in every 0.1 seconds. VPS can be set to any value however, only values of 1, 10, 100, and 1000 will have any effect on the program. For values inbetween, VPS will asume the next highest value (except VPS > 1000 which will default back to 1000).

**Excel File Name:**

By default values per second appears as follows within the program:

```
excel_file_name = 'Sensor_logs.xlsx'
```

This parameter can be set to anything so long as it ends in ''.xlsx'. The program will automatically export an excel file with the assigined name. 

NOTE: The program will not overwirte prexisting excel files meaning if an excel file of the same name already exists **it will not change!**

**Name(s) of Log File(s):**

By default values per second appears as follows within the program:

```
logfile_names = ['p0.log', 'p1.log', 'p2.log', 'p3.log', 'p4.log', 'p5.log', 'p6.log', 'p7.log', 'p8.log', 'p9.log']
```

This parameter is a list containing all the log files to be read (**IN ORDER FROM FIRST TO LAST**).

NOTE: If log file does not exist, the script will fail and return an error.
