import time
import serial

ser = serial.Serial('COM3',9600)

while 1:
    data = ser.readline()
    dcoded_data = data.decode('utf-8')
    usr_data = dcoded_data.split()[1]
    print (usr_data)