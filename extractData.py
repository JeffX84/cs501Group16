# -*- coding: utf-8 -*-
"""
Created on Sun Sep 23 16:15:09 2018

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

att_time = []
att_time_boot_ms = []
att_pitch = []
att_roll = []
att_yaw = []
att_p = []
att_q = []
att_r = []
imu_time = []
imu_xacc = []
imu_yacc = []
imu_zacc = []

for row in range(1,maxRow+1):
    if(sheet.cell(row=row,column=10).value == "mavlink_attitude_t"):
        att_time.append(sheet.cell(row=row,column=1).value)
        att_time_boot_ms.append(sheet.cell(row=row,column=12).value)
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
        
        
        
#Create new worksheet and output data
from openpyxl import Workbook
book = Workbook()
ws_att = book.active
ws_att.title = "ATTITUDE"
ws_imu = book.create_sheet("IMU")

#create headers
ws_att.cell(row=1,column=1).value = "attitude time"
ws_att.cell(row=1,column=1).value = "time boot ms"
ws_att.cell(row=1,column=2).value = "roll angle"
ws_att.cell(row=1,column=3).value = "pitch angle"
ws_att.cell(row=1,column=4).value = "yaw angle"
ws_att.cell(row=1,column=5).value = "roll rate"
ws_att.cell(row=1,column=6).value = "pitch rate"
ws_att.cell(row=1,column=7).value = "yaw rate"

ws_imu.cell(row=1,column=1).value = "IMU time"
ws_imu.cell(row=1,column=2).value = "IMU XAccel"
ws_imu.cell(row=1,column=3).value = "IMU YAccel"
ws_imu.cell(row=1,column=4).value = "IMU Zaccel"

for row in range(2,len(att_time)+2):
    dataindex = row-2
    ws_att.cell(row=row,column=1).value = att_time[dataindex]
    ws_att.cell(row=row,column=2).value = att_time_boot_ms[dataindex]    
    ws_att.cell(row=row,column=3).value = att_roll[dataindex]    
    ws_att.cell(row=row,column=4).value = att_pitch[dataindex]    
    ws_att.cell(row=row,column=5).value = att_yaw[dataindex]    
    ws_att.cell(row=row,column=6).value = att_p[dataindex]    
    ws_att.cell(row=row,column=7).value = att_q[dataindex]    
    ws_att.cell(row=row,column=8).value = att_r[dataindex]    
    
for row in range(2,len(imu_time)+2):
    dataindex = row-2
    ws_imu.cell(row=row,column=1).value = imu_time[dataindex]    
    ws_imu.cell(row=row,column=2).value = imu_xacc[dataindex]    
    ws_imu.cell(row=row,column=3).value = imu_yacc[dataindex] 
    ws_imu.cell(row=row,column=4).value = imu_zacc[dataindex] 
    
book.save("NewData.xlsx")
