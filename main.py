# This script automatically joins your Zoom classes using Python

# Import modules
import os
import pandas as pd
import pyautogui
import time
from datetime import datetime

# Open Zoom application
os.startfile(" ")

# Click the join meeting button
joinbtn=pyautogui.locateCenterOnScreen("join.png")
pyautogui.moveTo(joinbtn)
pyautogui.click()

# Enter meeting ID
meetingidbtn=pyautogui.locateCenterOnScreen("id.png")
pyautogui.moveTo(meetingidbtn)
pyautogui.write(meeting_id)

# Enter passcode
passcode=pyautogui.locateCenterOnScreen("")
pyautogui.moveTo(passcode)
pyautogui.write(password)

# Read Excel file with meeting IDs, passwords, and times
df = pd.read_excel('timings.xlsx',index=False)

# Loop to constantly check times and compare with Zoom meeting times
while True:
    #To get current time
    now = datetime.now().strftime("%H:%M")
    if now in str(df['Timings']):

        mylist=df["Timings"]
        mylist=[i.strftime("%H:%M") for i in mylist]
        c= [i for i in range(len(mylist)) if mylist[i]==now]
        row = df.loc[c]
        meeting_id = str(row.iloc[0,1])
        password= str(row.iloc[0,2])
        time.sleep(5)
        signIn(meeting_id, password)
        time.sleep(2)
        print('signed in')
        break

