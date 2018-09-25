"""
CS 50100 Group 16
Python script to extract raw drone data from Mission Planner output

openpyxl only works with *.xlsx (NOT CSV). Rename Book and Sheet to 'test' and save to xlsx
"""

import openpyxl as op
import numpy as np

book = op.load_workbook("test.xlsx")
sheet = book["test"]

#find number of rows
maxRow = sheet.max_row

print("%i number of rows found" % (maxRow))

#ahrs data
ahrs2_time, ahrs2_roll, ahrs2_pitch, ahrs2_yaw, ahrs2_altitude, ahrs2_lat, ahrs2_lng = [],[],[],[],[],[],[]
ahrs2_header = ["ahrs2_time","ahrs2 roll","ahrs2_pitch","ahrs2_yaw","ahrs2_altitude","ahrs2_lat","ahrs2_lng"]

ahrs3_time, ahrs3_roll, ahrs3_pitch, ahrs3_yaw, ahrs3_altitude, ahrs3_lat, ahrs3_lng = [],[],[],[],[],[],[]
ahrs3_header = ["ahrs3_time","ahrs3 roll","ahrs3_pitch","ahrs3_yaw","ahrs3_altitude","ahrs3_lat","ahrs3_lng"]

#attitude data
att_time, att_time_boot_ms, att_pitch, att_roll, att_yaw, att_p, att_q, att_r = [],[],[],[],[],[],[],[]
att_header = ["attitude time","time boot ms","roll angle","pitch angle","yaw angle","roll rate","pitch rate","yaw rate"]

#raw imu data
imu_time, imu_xacc, imu_yacc, imu_zacc, imu_xgyro, imu_ygyro, imu_zgyro, imu_xmag, imu_ymag, imu_zmag = [],[],[],[],[],[],[],[],[],[]
imu_header = ["IMU_time","Xaccel","Yaccel","Zaccel","XGyro","YGyro","ZGyro","XMag","YMag","ZMag"]

for row in range(1,maxRow+1):
    if(sheet.cell(row=row,column=10).value == "mavlink_ahrs2_t"):
        ahrs2_time.append(sheet.cell(row=row,column=1).value)
        ahrs2_roll.append(sheet.cell(row=row,column=12).value)
        ahrs2_pitch.append(sheet.cell(row=row,column=14).value)
        ahrs2_yaw.append(sheet.cell(row=row,column=16).value)
        ahrs2_altitude.append(sheet.cell(row=row,column=18).value)
        ahrs2_lat.append(sheet.cell(row=row,column=20).value)        
        ahrs2_lng.append(sheet.cell(row=row,column=22).value)
    if(sheet.cell(row=row,column=10).value == "mavlink_ahrs3_t"):      
        ahrs3_time.append(sheet.cell(row=row,column=1).value)
        ahrs3_roll.append(sheet.cell(row=row,column=12).value)
        ahrs3_pitch.append(sheet.cell(row=row,column=14).value)
        ahrs3_yaw.append(sheet.cell(row=row,column=16).value)
        ahrs3_altitude.append(sheet.cell(row=row,column=18).value)
        ahrs3_lat.append(sheet.cell(row=row,column=20).value)        
        ahrs3_lng.append(sheet.cell(row=row,column=22).value)
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
        imu_xgyro.append(sheet.cell(row=row,column=20).value)
        imu_ygyro.append(sheet.cell(row=row,column=22).value)
        imu_zgyro.append(sheet.cell(row=row,column=24).value)  
        imu_xmag.append(sheet.cell(row=row,column=26).value)
        imu_ymag.append(sheet.cell(row=row,column=28).value)
        imu_zmag.append(sheet.cell(row=row,column=30).value)          
        
        
#Create new worksheet and output data
from openpyxl import Workbook
book = Workbook()
ws_att = book.active
ws_att.title = "ATTITUDE"
ws_ahrs2 = book.create_sheet("AHRS_2")
ws_ahrs3 = book.create_sheet("AHRS_3")
ws_imu = book.create_sheet("IMU")

#create headers

for i in range(len(ahrs2_header)):
    ws_ahrs2.cell(row=1,column=i+1).value = ahrs2_header[i]
for row in range(2,len(ahrs2_time)+2):
    dataindex = row-2
    ws_ahrs2.cell(row=row,column=1).value = ahrs2_time[dataindex]
    ws_ahrs2.cell(row=row,column=2).value = ahrs2_roll[dataindex]
    ws_ahrs2.cell(row=row,column=3).value = ahrs2_pitch[dataindex]
    ws_ahrs2.cell(row=row,column=4).value = ahrs2_yaw[dataindex]
    ws_ahrs2.cell(row=row,column=5).value = ahrs2_altitude[dataindex]
    ws_ahrs2.cell(row=row,column=6).value = ahrs2_lat[dataindex]
    ws_ahrs2.cell(row=row,column=7).value = ahrs2_lng[dataindex]
    
for i in range(len(ahrs3_header)):
    ws_ahrs3.cell(row=1,column=i+1).value = ahrs3_header[i]
for row in range(2,len(ahrs2_time)+2):
    dataindex = row-2
    ws_ahrs3.cell(row=row,column=1).value = ahrs3_time[dataindex]
    ws_ahrs3.cell(row=row,column=2).value = ahrs3_roll[dataindex]
    ws_ahrs3.cell(row=row,column=3).value = ahrs3_pitch[dataindex]
    ws_ahrs3.cell(row=row,column=4).value = ahrs3_yaw[dataindex]
    ws_ahrs3.cell(row=row,column=5).value = ahrs3_altitude[dataindex]
    ws_ahrs3.cell(row=row,column=6).value = ahrs3_lat[dataindex]
    ws_ahrs3.cell(row=row,column=7).value = ahrs3_lng[dataindex]  

for i in range(len(att_header)):
    ws_att.cell(row=1,column=i+1).value = att_header[i]
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

for i in range(len(imu_header)):
    ws_imu.cell(row=1,column=i+1).value = imu_header[i]    
for row in range(2,len(imu_time)+2):
    dataindex = row-2
    ws_imu.cell(row=row,column=1).value = imu_time[dataindex]    
    ws_imu.cell(row=row,column=2).value = imu_xacc[dataindex]    
    ws_imu.cell(row=row,column=3).value = imu_yacc[dataindex] 
    ws_imu.cell(row=row,column=4).value = imu_zacc[dataindex] 
    ws_imu.cell(row=row,column=5).value = imu_xgyro[dataindex]    
    ws_imu.cell(row=row,column=6).value = imu_ygyro[dataindex] 
    ws_imu.cell(row=row,column=7).value = imu_zgyro[dataindex] 
    ws_imu.cell(row=row,column=8).value = imu_xmag[dataindex]    
    ws_imu.cell(row=row,column=9).value = imu_ymag[dataindex] 
    ws_imu.cell(row=row,column=10).value = imu_zmag[dataindex] 
    
book.save("NewData.xlsx")
