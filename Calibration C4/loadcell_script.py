import openpyxl
import sys
import os

"""
Openpyxl is a requirement for this script. Install by typing the following in CMD and pressing 
enter:

    > pip install openpyxl

Note: If using values per second (VPS) of 1000, there will be gaps in the excel file
due to diffrent sensors returnin values of slightly diffrent times. 

Only VPS values of 1, 10, 100 and 1000 can be used, values in between will result in using 
the next highest valid VPS value.

"""
# Change Parameters here 
vps = 10
excel_file_name = 'Sensor_logs.xlsx'
logfile_names = ['p0.log', 'p1.log', 'p2.log', 'p3.log', 'p4.log', 'p5.log', 'p6.log', 'p7.log'] #Insert in order

#########################################################################################
# Script Begins here
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Sensor Log Ouput"

# Variable declaration for use below
prev_time = 'null'
prev_line = 1

time_column = "A{}"
col_dict = {'c0': 'B{}', 
            'c1': 'C{}', 
            'c2': 'D{}', 
            'c3': 'E{}', 
            'c4': 'F{}', 
            'c5': 'G{}', 
            'c6': 'H{}', 
            'c7': 'I{}'}

# Sets up column titles. Note: time is in hours:minutes:seconds,milliseconds
sheet['A1'] = "Time"

for i in range(8):
    col = 'c{}'.format(i)
    sheet[col_dict[col].format(1)] = col.upper()

# Opens File(s)
for file in logfile_names:
    f = open(os.path.join(sys.path[0], file), 'r')
    log_lines = f.readlines()

    for i in range(len(log_lines)):
        """
        Reads log file data based on following format: 
        
        YYY-MM-DD HH:MM:SS,MSS INFO [sensorValueLogger] cx,value
        
        Wrong spacing and syntax withing logfile will likely result in error)

        """
        # Splits up data for usage
        current_row_split = log_lines[i].split()
        sensorvalue = current_row_split[4].split(',')
        time = current_row_split[1].split(',')

        current_time = float(time[1])

        # Rounds milliseconds acording to desired sensor ouput rate in excel file
        if 10 < vps <= 100:
            current_time = round((current_time/1000), 2) * 1000
        elif 1 < vps <= 10:
            current_time = round((current_time)/1000, 1) * 1000
        elif vps == 1:
            current_time = 0


        # Writes Data to excel file
        time_string = time[0] + ',' + str(int(current_time))
        if prev_time != time_string:
            sheet[time_column.format(prev_line + 1)] = (time_string)
            sheet[(col_dict[sensorvalue[0]]).format(prev_line + 1)] = sensorvalue[1]
            prev_time = time_string
            prev_line = prev_line + 1
            #print((col_dict[sensorvalue[0]]).format(i + 2))
        else:
            sheet[(col_dict[sensorvalue[0]]).format(prev_line)] = sensorvalue[1]

    f.close()

# Saves excel file
wb.save(filename = excel_file_name)