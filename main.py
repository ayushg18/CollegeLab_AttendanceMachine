import cv2
from openpyxl import load_workbook
from pyzbar.pyzbar import *
import pandas as pd
import datetime
from time import strftime
import os
import pyttsx3
from styleframe import StyleFrame

vid = cv2.VideoCapture(0)

def barcode():
    while True:
        ret, img = vid.read()
        for barcode in decode(img):
            dataUID = barcode.data.decode('utf-8')
            print(dataUID)
            return(dataUID)

        cv2.imshow('Camera', img)
        cv2.waitKey(10)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def unauthorisedOutLog():       #Function for printing out log
    df = pd.read_excel(unauthorisedLogFile)
    index = df.index
    maxRowOfLog = len(index)
    i=0
    while i<maxRowOfLog:
        if( uid == str(df['UID'][i]) and ('-' + uid + '-') == str(df['OUT Time'][i])  ):
            a = ('-' + uid + '-')
            b = str(timeOfScan)
            
            df = (df.replace(to_replace=a, value=b))
            df.to_excel(unauthorisedLogFile, sheet_name='Sheet_1', index = False)
            speak('Out')
            return 5
        i+=1


def unauthorisedLog():       #Function for printing IN log
    with open(unauthorisedLogFile, mode='a'):
        old_df = pd.read_excel(unauthorisedLogFile, engine='openpyxl')
        inputData = [[uid, timeOfScan, ('-' + uid + '-')]]
        new_df = pd.DataFrame(inputData, columns=['UID', 'IN Time', 'OUT Time'])

        frames = [old_df, new_df]
        result = pd.concat(frames)
        result.to_excel(unauthorisedLogFile, sheet_name='Sheet_1', index = False)
        speak('In')


def unauthorised():       #Function to check file were already exists or not for unauthorised member
    if(os.path.exists(f'./{unauthorisedLogFile}') == True):
        condition = unauthorisedOutLog()
        if(condition != 5):
            unauthorisedLog()

    elif(os.path.exists(f'./{unauthorisedLogFile}') == False):
        with open(unauthorisedLogFile, 'a'):
            inputData = [[uid, timeOfScan, ('-' + uid + '-')]]
            df = pd.DataFrame(inputData, columns=['UID', 'IN Time', 'OUT Time'])
            df.to_excel(unauthorisedLogFile, index=False)
            speak('In')


def scanFromFile():       #Function to get data for database and return required data else return None
    df = pd.read_excel(r"uidDataExelSheet.xlsx")
    index = df.index
    maxRow = len(index)
    for i in range(0, maxRow):
        if(uid == str(df['UID'][i])):
            # print(df['UID'][i])
            # print(df['Semester'][i])
            # print(df['Name'][i])
            userUID = (df['UID'][i])
            userSemester = (df['Semester'][i])
            userName = (df['Name'][i])
            return (userUID, userName, userSemester)


def storeLog():       #Function to store data if file were already exits
    with open(logExcelFile, mode='a'):
        old_df = pd.read_excel(logExcelFile, engine='openpyxl')
        inputData = [[dataGetFromScanning[0], dataGetFromScanning[1], dataGetFromScanning[2], timeOfScan, ('-' + dataGetFromScanning[0] + '-')]]
        new_df = pd.DataFrame(inputData, columns=['UID', 'Name', 'Semester', 'IN Time', 'OUT Time'])
        frames = [old_df, new_df]
        result = pd.concat(frames)
        result.to_excel(logExcelFile, sheet_name='Sheet_1', index = False)


def scanLogFile():       #Function for OUT for authorised
        df = pd.read_excel(logExcelFile)
        index = df.index
        maxRowOfLog = len(index)
        for i in range(0, maxRowOfLog):
            if( uid == str(df['UID'][i]) and ('-' + uid + '-') == str(df['OUT Time'][i])  ):
                a = (df['OUT Time'][i])
                b = str(timeOfScan)

                df = (df.replace(to_replace=a, value=b))
                df.to_excel(logExcelFile, sheet_name='Sheet_1', index = False)
                speak('Out')
                return 5


# --> Main Body of Program <--

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voices', voices[0].id)

speak('Attendance Machine is ON')

# def mainBody():
while True:
    uid = barcode()

    filename = datetime.datetime.now()
    logExcelFile = (filename.strftime("%d - %B - %Y")+".xlsx")   # %d - date, %B - month, %Y - Year     
    # This is Authorisied File

    unauthorisedLogFile = (filename.strftime("Unauthorised %d - %B - %Y")+".xlsx")      # This is for Unauthorisied File


    timeOfScan = strftime('%H : %M : %S : %p')      # %H - Hours, %M - Minutes, %S - Seconds, %p - AM or PM
    # But their is no need of %p - AM or PM because of hours system is in 24 hours.

    dataGetFromScanning = scanFromFile()      # Storing return value and also calling a function
    # print(dataGetFromScanning)

    if(dataGetFromScanning != None):
        if (os.path.exists(f'./{logExcelFile}') == True):
            condition = scanLogFile()
            # Here '5' is just for flag, instead of '5' we can use any number
            if(condition != 5):
                storeLog()
                speak('In')

        elif (os.path.exists(f'./{logExcelFile}') == False):
            # For File/Directory --> os.path.exists('./filename with format')
            # For Files --> os.path.isfile('./filename with format')
            # For Directory --> os.path.isdir('./filename with format')

            with open(logExcelFile, 'a'):
                inputData = [[dataGetFromScanning[0], dataGetFromScanning[1], dataGetFromScanning[2], timeOfScan, ('-' + dataGetFromScanning[0] + '-')]]
                df = pd.DataFrame(inputData, columns=['UID', 'Name', 'Semester', 'IN Time', 'OUT Time'])
                df.to_excel(logExcelFile, index=False)
                speak('In')

    elif(dataGetFromScanning == None):
        unauthorised()