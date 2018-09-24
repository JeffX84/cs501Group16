"""
CS 50100 Group 16
Python script to extract raw drone data from Mission Planner output

openpyxl only works with *.xlsx (NOT CSV). Rename Book and Sheet to 'test' and save to xlsx
"""

import openpyxl as op

book = op.load_workbook("test.xlsx")
sheet = book["test"]

#find number of rows
maxRow = sheet.max_row

print("%i number of rows found" % (maxRow))

#declare local lists
att_time = att_pitch = att_roll = att_yaw = att_p = att_q = att_r = [] # attitude data
imu_time = imu_xacc = imu_yacc = imu_zacc = [] # raw_imu data

for row in range(1,maxRow+1):
    if(sheet.cell(row=row,column=10).value == "mavlink_attitude_t"):
        att_time.append(sheet.cell(row=row,column=1).value)
        att_roll.append(sheet.cell(row=row,column=14).value)
        att_pitch.append(sheet.cell(row=row,column=16).value)
        att_yaw.append(sheet.cell(row=row,column=18).value)        
        att_p.append(sheet.cell(row=row,column=20).value)
        att_q.append(sheet.cell(row=row,column=22).value)
        att_r.append(sheet.cell(row=row,column=24).value)    
    if(sheet.cell(row=row,column=10).value == "mavlink_raw_imu_t"):
        imu_time.append(sheet.cell(row=row,column=1).value)
        imu_xacc.append(sheet.cell(row=row,column=14).value)
        imu_yacc.append(sheet.cell(row=row,column=16).value)
        imu_zacc.append(sheet.cell(row=row,column=18).value)        
        
        
        
#Create new data
from openpyxl import Workbook
book = Workbook()
sheet = book.active

#create headers
sheet.cell(row=1,column=1).value = "attitude time"
sheet.cell(row=1,column=2).value = "roll angle"
sheet.cell(row=1,column=3).value = "pitch angle"
sheet.cell(row=1,column=4).value = "yaw angle"
sheet.cell(row=1,column=5).value = "roll rate"
sheet.cell(row=1,column=6).value = "pitch rate"
sheet.cell(row=1,column=7).value = "yaw rate"

sheet.cell(row=1,column=8).value = "IMU time"
sheet.cell(row=1,column=9).value = "IMU XAccel"
sheet.cell(row=1,column=10).value = "IMU YAccel"
sheet.cell(row=1,column=11).value = "IMU Zaccel"

for row in range(2,len(att_time)+2):
    dataindex = row-2
    sheet.cell(row=row,column=1).value = att_time[dataindex]
    sheet.cell(row=row,column=2).value = att_roll[dataindex]    
    sheet.cell(row=row,column=3).value = att_pitch[dataindex]    
    sheet.cell(row=row,column=4).value = att_yaw[dataindex]    
    sheet.cell(row=row,column=5).value = att_p[dataindex]    
    sheet.cell(row=row,column=6).value = att_q[dataindex]    
    sheet.cell(row=row,column=7).value = att_r[dataindex]    
    
for row in range(2,len(imu_time)+2):
    dataindex = row-2
    sheet.cell(row=row,column=8).value = imu_time[dataindex]    
    sheet.cell(row=row,column=9).value = imu_xacc[dataindex]    
    sheet.cell(row=row,column=10).value = imu_yacc[dataindex] 
    sheet.cell(row=row,column=11).value = imu_zacc[dataindex] 
    
book.save("NewData.xlsx")
