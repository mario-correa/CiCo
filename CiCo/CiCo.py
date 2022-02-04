#for the xls
import pandas as pd

#to get the time
from datetime import datetime
import time

#for the fingerprint sensor
import adafruit_fingerprint
import serial

#for the printer
from escpos.printer import Network

#for the buzzer
import RPi.GPIO as GPIO

#-------------------------------------------------------------------------
#code from the adafruit example
def get_fingerprint():
    #Get a finger print image, template it, and see if it matches!
    #print("Waiting for image...")
    while finger.get_image() != adafruit_fingerprint.OK:
        pass
    #print("Templating...")
    if finger.image_2_tz(1) != adafruit_fingerprint.OK:
        return False
    #print("Searching...")
    if finger.finger_search() != adafruit_fingerprint.OK:
        return False
    return True
#-------------------------------------------------------------------------
#make a sound..
def beep(x=1):
    for i in range(x):
        GPIO.output(buzzer,GPIO.HIGH)
        time.sleep(0.03)
        GPIO.output(buzzer,GPIO.LOW)
#-------------------------------------------------------------------------
#config for the files with employee information

people = '/home/pi/CiCo/people.xls' #xls maps names to sensor positions
df_people = pd.read_excel(people)

#config for the buzzer
GPIO.setmode(GPIO.BCM)
buzzer = 21 #the hardware conection
GPIO.setup(buzzer, GPIO.OUT)

#config for the output files
reportFolder = '/home/pi/CiCo Reports/'

#config for the printer
ip = "192.168.86.54" #Printer IP Address

#config for the fingerprint sensor
uart = serial.Serial("/dev/ttyAMA0", baudrate=57600, timeout=1)
finger = adafruit_fingerprint.Adafruit_Fingerprint(uart)

#get the header info
h = open('/home/pi/CiCo/header.txt','r')
header = h.read()
h.close()
header = header+"\n\n"


if finger.read_templates() != adafruit_fingerprint.OK:
        beep(3)
        raise RuntimeError("Failed to read templates")

#main loop, so it will be on 24x7
while(True):
    if get_fingerprint():
        beep()
        person = df_people[df_people['Pos']==finger.finger_id]['Name'].values[0]
        now = datetime.now()
        current_time = now.strftime("%H:%M:%S")
        current_date = now.strftime("%d/%m/%Y")

        #Change the output file
        df = pd.read_excel(reportFolder+person+'.xls')
        log = [person,current_date,current_time]
        a_series = pd.Series(log, index = df.columns)
        df = df.append(a_series, ignore_index=True)
        
        df.to_excel(reportFolder+person+'.xls', index=False)

        try:
            #Print the recipt
            p = Network(ip,timeout=20)
            p.text(header)
            p.text("Nome: "+person+"\n")
            p.text("Data: " + current_date)
            p.text("\n")
            p.text("Hora: " + current_time)
            p.text("\n")
        except:
            pass
        finally:
            p.cut()
            p.close()