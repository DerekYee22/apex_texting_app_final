import pandas as pd
from datetime import datetime, timedelta 
from twilio.rest import Client
import subprocess
import shutil
import os
import re
from tkinter import messagebox

# insert base path here (removed for privacy purposes)
base_path = ''
logPath = os.path.join(base_path,'log.csv')
outlookPath = os.path.join(base_path,'calendar.csv')

finishedOutlookPath = os.path.join(base_path,'Finished Outlook Files')

secrets = pd.read_csv(os.path.join(base_path, 'secrets.csv'))

account_sid = str(secrets.loc[0, 'account_sid']).strip()
auth_token = str(secrets.loc[0, 'auth_token']).strip()
my_twilio = str(secrets.loc[0, 'my_twilio']).strip()
apex_phone = str(secrets.loc[0, 'apex_phone']).strip()
DATE = datetime.now()

def extract_phone(body):
    if pd.isna(body):
        return None
    match = re.search(r'(\d{10})', str(body))
    return match.group(1) if match else None

def confirmText(appointments):
    text = ""
    for appointment in appointments:
        text += f"Patient Name: {appointment['patientFirstName']} {appointment['patientLastName']}\n"
        text += f"Phone Number: {appointment['phoneNumber']}\n"
        text += f"Appointment Date: {appointment['date']}\n"
        text += f"Appointment Time: {appointment['time']}\n\n"
    root = messagebox.askyesno("Confirm SMS", f"{text}Are you sure you want to send the text messages?")
    if root:
        return True
    else:
        return False
    
def getInfo(i):
    body = calendar.iloc[i, 4]
    phone = extract_phone(body)
    # Skip if no phone number
    if not phone:
        return None

    startString = str(calendar.iloc[i, 1]).strip()
    startDateTime = datetime.strptime(startString, "%Y-%m-%d %H:%M:%S.%f")

    dateString = startDateTime.strftime("%m/%d/%y")
    timeString = startDateTime.strftime("%I:%M %p").lstrip("0")

    patientName = str(calendar.iloc[i, 0]).strip()
    if patientName.startswith("*"):
        patientName = patientName[1:].strip()
    patientNameArray = patientName.split(",")
    if len(patientNameArray) == 2:
        patientLastName = patientNameArray[0].strip()
        patientFirstName = patientNameArray[1].strip().split(" ")[0]
    else:
        name_parts = patientName.split(" ")
        patientFirstName = name_parts[1] if len(name_parts) > 1 else ""
        patientLastName = name_parts[0]

    phoneNumber = '1' + phone

    return {
        "phoneNumber": phoneNumber,
        "patientFirstName": patientFirstName,
        "patientLastName": patientLastName,
        "date": dateString,
        "time": timeString
    }

def sendSMS(appointment):
    message = '''Apex Physical Therapy Specialists: This is a friendly reminder of your appointment scheduled for %s at %s.
If you need to contact the clinc please call %s. \n\n DO NOT REPLY''' % (appointment["date"], appointment["time"], apex_phone)
    # send sms
    client = Client(account_sid, auth_token)

    sms_message = client.messages.create(
        to=appointment["phoneNumber"], 
        from_=my_twilio,
        body=message)
    
    return sms_message.status

def writeLog(appointment):
    with open(logPath, "a") as file:
        new_entry = (
            f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
            f"Phone Number: {appointment['phoneNumber']}\n"
            f"Patient Appointment Time: {appointment['time']}\n"
            f"Patient Name: {appointment['patientFirstName']} {appointment['patientLastName']}\n"
            f"Status: {sendSMS(appointment)} \n\n"
        )
    if os.path.exists(logPath):
            with open(logPath, "r", encoding="utf-8") as file:
                existing_log = file.read()
    else:
        existing_log = ""

    with open(logPath, "w") as file:
        file.write(new_entry + existing_log)


def text(date, calendar_name):
    global calendar
    calendar = pd.read_csv(outlookPath)
    appointmentsToday = []
    i = 1
    while i < len(calendar):
        info = getInfo(i)
        if info is not None:
            appointmentsToday.append(info)
        i += 1

    if len(appointmentsToday) == 0:
        messagebox.showinfo("Failure", f"No appointments found for {date} in {calendar_name}.")
        return 0
    
    if not (confirmText(appointmentsToday)):
        messagebox.showinfo("Failure", f"Cancelled sending SMS")
        return 0
    
    for appointment in appointmentsToday:
        writeLog(appointment)
    
    with open(logPath, "r") as file:
        existingLog = []
        for line in file:
            existingLog.append(line)

    # deletes logs that are 1 week old
    logIndex = 0
    while logIndex < len(existingLog):
        if existingLog[logIndex][0:4] == "Date":
            sendDate = datetime.strptime(existingLog[logIndex][6:-2], '%Y-%m-%d %H:%M')
            oneWeekAgo = DATE - timedelta(weeks=1)
            if sendDate < oneWeekAgo:
                existingLog = existingLog[:logIndex]
            else:
                logIndex += 1
        else:
            logIndex += 1

    # Write the updated log back to the file
    with open(logPath, "w") as file:
        file.writelines(existingLog)

    # moves outlook file to Finished Outlook Files Folder
    if os.path.exists(os.path.join(finishedOutlookPath,'calendar.csv')):
        os.remove(os.path.join(finishedOutlookPath,'calendar.csv'))
    shutil.move(outlookPath, finishedOutlookPath)

    messagebox.showinfo("Sucess!", f"Text messages sent successfully!\nLog file created at {logPath}")

    # opens log
    subprocess.call(["notepad.exe", logPath], shell=True)