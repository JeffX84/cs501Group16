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
ahrs2_header = ["timestamp","ahrs2 roll","ahrs2_pitch","ahrs2_yaw","ahrs2_altitude","ahrs2_lat","ahrs2_lng"]

ahrs3_time, ahrs3_roll, ahrs3_pitch, ahrs3_yaw, ahrs3_altitude, ahrs3_lat, ahrs3_lng = [],[],[],[],[],[],[]
ahrs3_header = ["timestamp","ahrs3_roll","ahrs3_pitch","ahrs3_yaw","ahrs3_altitude","ahrs3_lat","ahrs3_lng"]

#attitude data
att_time, att_time_boot_ms, att_pitch, att_roll, att_yaw, att_p, att_q, att_r = [],[],[],[],[],[],[],[]
att_header = ["attitude timestamp","time boot ms","roll angle","pitch angle","yaw angle","roll rate","pitch rate","yaw rate"]

#Battery data
batt_time, batt_cc, batt_ec, batt_temp, batt_curr, batt_func, batt_rem,  batt_time_rem = [],[],[],[],[],[],[],[]
batt_header = ["battery timestamp", "Current Consumed","Energy Consumed","Battery temperature","battery current","battery function","battery remaining","battery time remaining"]

#raw imu data
imu_time, imu_xacc, imu_yacc, imu_zacc, imu_xgyro, imu_ygyro, imu_zgyro, imu_xmag, imu_ymag, imu_zmag = [],[],[],[],[],[],[],[],[],[]
imu_header = ["IMU timestamp","Xaccel","Yaccel","Zaccel","XGyro","YGyro","ZGyro","XMag","YMag","ZMag"]

#servo data
servo_time, servo_utim, servo1_raw, servo2_raw, servo3_raw, servo4_raw = [],[],[],[],[],[]
servo_header = ["Servo timestamp","Servo utime","Servo 1","Servo 2","Servo 3","Servo 4"]

#vibration data
vib_time, vib_x,vib_y,vib_z = [],[],[],[]
vib_header = ["Vibration timestamp","vibration_x","vibration_y","vibration_z"]

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
    if(sheet.cell(row=row,column=10).value == "mavlink_battery_status_t"):     
        batt_time.append(sheet.cell(row=row,column=1).value)     
        batt_cc.append(sheet.cell(row=row,column=12).value)     
        batt_ec.append(sheet.cell(row=row,column=14).value)     
        batt_temp.append(sheet.cell(row=row,column=16).value)     
        batt_curr.append(sheet.cell(row=row,column=20).value)     
        batt_func.append(sheet.cell(row=row,column=24).value)     
        batt_rem.append(sheet.cell(row=row,column=28).value)     
        batt_time_rem.append(sheet.cell(row=row,column=30).value)     
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
    if(sheet.cell(row=row,column=10).value == "mavlink_servo_output_raw_t"):
        servo_time.append(sheet.cell(row=row,column=1).value)
        servo_utim.append(sheet.cell(row=row,column=12).value)
        servo1_raw.append(sheet.cell(row=row,column=14).value)
        servo2_raw.append(sheet.cell(row=row,column=16).value)
        servo3_raw.append(sheet.cell(row=row,column=18).value)
        servo4_raw.append(sheet.cell(row=row,column=20).value)
    if(sheet.cell(row=row,column=10).value == "mavlink_vibration_t"):
        vib_time.append(sheet.cell(row=row,column=1).value)
        vib_x.append(sheet.cell(row=row,column=14).value)
        vib_y.append(sheet.cell(row=row,column=16).value)
        vib_z.append(sheet.cell(row=row,column=18).value)
 

#Create new worksheet and output data
from openpyxl import Workbook
book = Workbook()
ws_att = book.active
ws_att.title = "ATTITUDE"
ws_ahrs2 = book.create_sheet("AHRS_2")
ws_ahrs3 = book.create_sheet("AHRS_3")
ws_batt = book.create_sheet("BATTERY")
ws_imu = book.create_sheet("IMU")
ws_servo = book.create_sheet("SERVO")
ws_vib = book.create_sheet("VIBRATION")

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
for row in range(2,len(ahrs3_time)+2):
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

for i in range(len(batt_header)):
    ws_batt.cell(row=1,column=i+1).value = batt_header[i]    
for row in range(2,len(batt_time)+2):
    dataindex = row-2
    ws_batt.cell(row=row,column=1).value = batt_time[dataindex]  
    ws_batt.cell(row=row,column=2).value = batt_cc[dataindex]  
    ws_batt.cell(row=row,column=3).value = batt_ec[dataindex]  
    ws_batt.cell(row=row,column=4).value = batt_temp[dataindex]  
    ws_batt.cell(row=row,column=5).value = batt_curr[dataindex]  
    ws_batt.cell(row=row,column=6).value = batt_func[dataindex]  
    ws_batt.cell(row=row,column=7).value = batt_rem[dataindex]  
    ws_batt.cell(row=row,column=8).value = batt_time_rem[dataindex]  
  
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
    
for i in range(len(servo_header)):
    ws_servo.cell(row=1,column=i+1).value = servo_header[i]    
for row in range(2,len(servo_time)+2):
    dataindex = row-2
    ws_servo.cell(row=row,column=1).value = servo_time[dataindex]        
    ws_servo.cell(row=row,column=2).value = servo_utim[dataindex]  
    ws_servo.cell(row=row,column=3).value = servo1_raw[dataindex]  
    ws_servo.cell(row=row,column=4).value = servo2_raw[dataindex]  
    ws_servo.cell(row=row,column=5).value = servo3_raw[dataindex]  
    ws_servo.cell(row=row,column=6).value = servo4_raw[dataindex]  
    
for i in range(len(vib_header)):
    ws_vib.cell(row=1,column=i+1).value = vib_header[i]    
for row in range(2,len(vib_time)+2):
    dataindex = row-2
    ws_vib.cell(row=row,column=1).value = vib_time[dataindex]
    ws_vib.cell(row=row,column=2).value = vib_x[dataindex]      
    ws_vib.cell(row=row,column=3).value = vib_y[dataindex]      
    ws_vib.cell(row=row,column=4).value = vib_z[dataindex]      
    
book.save("NewData.xlsx")
