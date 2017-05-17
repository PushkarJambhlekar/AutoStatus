""" Status mail bot docstring
    Sends status
"""
import win32com.client as win32
import subprocess
import os, sys
import uuid
import datetime
import time
import docMap

def sendStatus():
    now = datetime.datetime.now()
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = '***'
    mail.Subject = 'Message subject'
    mail.Subject = '[Status] Pushkar Jambhlekar '+now.strftime("%Y-%m-%d")
    mail.HTMLBody = docMap.GetStatus()
    mail.Send()

statusSent = 0
def sendStatusOnMonday():
    global statusSent
    now = datetime.datetime.now()
    day = now.strftime("%A")
    if day == 'Sunday' and statusSent == 0:
        sendStatus()
        statusSent = 1
    statusSent = 0

def do_work():
    sendStatusOnMonday()

sendStatus()

while 1:
    do_work()
    time.sleep(3600)
