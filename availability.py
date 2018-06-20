import openpyxl
import json
from datetime import datetime
import calendar
from Schedule import scheduling
import time
import serial



current_time=datetime.now()

workbook = openpyxl.load_workbook('Schedule.xlsx')
worksheet = workbook[calendar.day_name[current_time.weekday()]]
def reload():
  """ reset all data"""	
  outfile=open('availability.json', mode='w')
  status_upload={}
  for row in worksheet.iter_rows(min_row=2, min_col=0, max_col=1):
      for cell in row:
          status_upload[str(cell.value)]="absent"
  json.dump(status_upload, outfile)
  outfile.close()


def modify(index):
  """
  enter index of the record to be changed
  """	
  file_read=open("availability.json",mode="r")	
  check_availability=json.load(file_read) 
  if check_availability[str(index)]=="present":
      check_availability[str(index)]="absent"
  elif check_availability[str(index)]=="absent":
      check_availability[str(index)]="present"
  file_read.close()
  file_write=open("availability.json",mode="w+")	  
  json.dump(check_availability,file_write)
  file_write.close()
  scheduling()

reload()
ser = serial.Serial('COM3',9600)

while 1:
    data = ser.readline()
    dcoded_data = data.decode('utf-8')
    usr_data = dcoded_data.split()[1]
    print("Input ID: " + usr_data )
    modify(usr_data)
    




