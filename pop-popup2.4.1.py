# pop-popup.py
# Mike Tsak 20/8/22
version = "2.4.1"
#
# requarements:  python -m pip install pysimplegui
#               pip3 install gspread
#               pip3 install --upgrade google-api-python-client oauth2client
#               pip install google-auth
#               pip install manager
#               pip install pandas

#import PySimpleGUIQt as sg
from PySimpleGUIQt import Window
from PySimpleGUIQt import WIN_CLOSED
from PySimpleGUIQt import Text
from PySimpleGUIQt import Button
from PySimpleGUIQt import InputText
from PySimpleGUIQt import theme
from PySimpleGUIQt import Popup
from PySimpleGUIQt import Checkbox

from socket import gethostname

import time

import keyboard

from gspread import service_account_from_dict
from google.oauth2.service_account import Credentials

from datetime import datetime


# sa = gspread.service_account(filename="credentialsfile.json") #filename = "service_account.jprg"
# sh = sa.open("testapi")


def bootlogger():
    sheet1 = google_sh.worksheet("bootlogger")
    chells = sheet1.acell('d2').value
    chell = int(chells)
    chell += 1
    sheet1.update('d2', chell)
    dt = datetime.now()
    dt = str(dt)
    sheet1.update_cell(chell, 1, dt)
    sheet1.update_cell(chell, 2, gethostname())
    sheet1.update_cell(chell, 3, version)


def loginlogger():
    sheet1 = google_sh.worksheet("loginlogger")
    chells = sheet1.acell('d2').value
    chell = int(chells)
    chell += 1
    sheet1.update('d2', chell)
    dt = datetime.now()
    dt = str(dt)
    sheet1.update_cell(chell, 1, dt)
    user_ids = str(user_id)
    sheet1.update_cell(chell, 2, user_ids)


def call_answerd():
    dt = datetime.now().date()
    dt = str(dt)
    sheet1.update_cell(row, coll, dt)
    tm = datetime.now().time()
    tm = str(tm)
    sheet1.update_cell(row, coll + 1, tm)


def call_not_answerd():
    dt = datetime.now().date()
    dt = str(dt)
    sheet1.update_cell(row, coll, dt)

def check_anapantiti(checkflag):
	if checkflag == True:
		sheet1.update_cell(row, coll + 4, "NoAnswer")
	else:
		mysecondmenu()


def myfirstmenu():
    tm = datetime.now().time()
    tm = str(tm)
    sheet1.update_cell(row, coll + 2, tm)

    layout1 = [[Text("Επέλεξε τι τηλεφώνημα ήταν:", font=font)], [Button("Εισερχόμενη 📲", font=font)],
               [Button("Εξερχόμενη 📞", font=font)],
               [Button("Εσωτερική 🏢", font=font)],
               [Checkbox("Δεν απάντησε📵", default=False, key="-IN-", font=font3)]]  # hard beded becuse og a pygui bag

    window1 = Window("Phone Tracking 1st menu", layout1, icon=r'C:\popicon.ico')

    while 1:
        event, values = window1.read()
        # End program if user closes window or
        # presses the OK button
        checkboxcheck = values["-IN-"]
        if event == WIN_CLOSED:
            window1.close()
            break
        if event == "Εισερχόμενη 📲":
            window1.close()
            sheet1.update_cell(row, coll + 3, "Εισ")
            check_anapantiti(checkboxcheck)
        if event == "Εξερχόμενη 📞":
            window1.close()
            sheet1.update_cell(row, coll + 3, "Εξερ")
            window1.close()
            check_anapantiti(checkboxcheck)
        if event == "Εσωτερική 🏢":
            window1.close()
            sheet1.update_cell(row, coll + 3, "Εσωτ")
            check_anapantiti(checkboxcheck)
        break


def mysecondmenu():
    layout2 = [[Text("Επέλεξε τι τύπου ήταν:", font=font)], [Button("Delay calls ⏰", font=font)],
               [Button("Missing delivery notes info 🛵", font=font)], [Button("Problem with the order 🛒", font=font)],
               [Button("Technical Issue 🛠", font=font)],
               [Button("Other 😶", font=font)]]  # hard beded becuse og a pygui bag
    window1 = Window("Phone Tracking 2nd menu", layout2, icon=r'C:\popicon.ico')

    while 1:
        event, values = window1.read()
        # End program if user closes window1 or
        # presses the OK button
        if event == "Exit❌" or event == WIN_CLOSED:
            window1.close()
        if event == "Delay calls ⏰":
            window1.close()
            sheet1.update_cell(row, coll + 4, "Delay")
        if event == "Missing delivery notes info 🛵":
            window1.close()
            sheet1.update_cell(row, coll + 4, "Note")
        if event == "Problem with the order 🛒":
            window1.close()
            sheet1.update_cell(row, coll + 4, "POrder")
        if event == "Technical Issue 🛠":
            window1.close()
            sheet1.update_cell(row, coll + 4, "Tech")
        if event == "Other 😶":
            window1.close()
            other_input()
        break


def other_input():
    sheet1.update_cell(row, coll + 4, "Other")

    layout3 = [[Text("Πληκτρολογήστε γιατί: (Optional)", font=font)], [InputText(font=font)],
               [Button("Submit", font=font)]]  # hard beded becuse og a pygui bag
    window2 = Window("Other", layout3, icon=r'C:\popicon.ico')

    while 1:
        event, values = window2.read()
        if event == WIN_CLOSED or event == 'Submit':
            # values = values['-IN-']
            values = str(values)
            sheet1.update_cell(row, coll + 5, values)
            window2.close()
            break
        window2.close()


def change_theme():
    layout5 = layout = [[Button('Αυτό που αρέσει στην user1(Μπλεδούλι)🔵', font=font)],
                        [Button('Pop🌸', font=font), Button('Λευκό⚪', font=font), Button('Μαύρο⚫', font=font), Button('Κατσάου!🏎️', font=font)]]
    windowt = Window("Theme Selection", layout5, icon=r'C:\popicon.ico')
    while 1:
        event, values = windowt.read()
        if event == WIN_CLOSED:
            windowt.close()
            break
        if event == 'Pop🌸':
            windowt.close()
            return ('DarkPurple5')
            break
        if event == 'Λευκό⚪':
            windowt.close()
            return ('LigthGray1')
            break
        if event == 'Μαύρο⚫':
            windowt.close()
            return ('DarkBlack1')
            break
        if event == 'Κατσάου!🏎️':
            windowt.close()
            return ('DarkRed2')
        if event == 'Αυτό που αρέσει στην user1(Μπλεδούλι)🔵':
            return ('LightBlue')
            break

def oldtheme():
    google_sh = gc.open("pop-popup debugging data")
    sheet1 = google_sh.worksheet("CC Agents")
    theme = sheet1.cell(user_id, 1).value
    if theme == None:
        return ('LightBlue')
    else:
        return (theme)

def setthemeindb():
    google_sh = gc.open("pop-popup debugging data")
    sheet1 = google_sh.worksheet("CC Agents")
    sheet1.update_cell(user_id, 1, theme1)

def get_from_chell():
    if FailSafe == 0:
        chells = sheet1.cell(1, coll).value
        chell = int(chells)
        num = chell
        chell += 1
        sheet1.update_cell(1, coll, chell)
        return (num)
    elif FailSafe == 1:
        num = 1
        while 1:
            val = sheet1.cell(num, coll).value
            if val == None:
                  break
            num += 1
        num -= 1
    return (num)


def get_from_name(nameid):
    if nameid == 105:  # user1
        num = 7
    elif nameid == 118:  # χρη
        num = 14
    elif nameid == 106:  # μελπο
        num = 21
    elif nameid == 102:  # user4
        num = 28
    elif nameid == 1955:  # winclosed
        num = 35
    else:
        Popup("CRITICAL ERROR", font=font)
    return (num)


def getname(nameid):
    sheet1 = google_sh.worksheet("CC Agents")
    if nameid == 105:  # user1
        name = sheet1.acell('d2').value
    elif nameid == 118:  # χρη
        name = sheet1.acell('d3').value
    elif nameid == 106:  # μελπο
        name = sheet1.acell('d4').value
    elif nameid == 102:  # user4
        name = sheet1.acell('d5').value
    elif nameid == 1955:
        name = "error"
    else:
        Popup("CRITICAL ERROR", font=font)
    return (name)

def greedgeexaltedones():
    currentTime = int(time.strftime('%H'))
    if currentTime < 12:
        Popup('Καλημέρα', name, "😉", font=font)
    elif 12 <= currentTime < 21:
        Popup('Καλησπέρα', name, "😉", font=font)
    else:
        Popup('Καλή βάρδια', name, "😉", font=font)

def bugreport():
    google_sh = gc.open("pop-popup debugging data")
    sheet1 = google_sh.worksheet("bug report")
    chells = sheet1.acell('f2').value
    chell = int(chells)
    chell += 1
    sheet1.update('f2', chell)
    dt = datetime.now()
    dt = str(dt)
    user_ids = str(user_id)
    sheet1.update_cell(chell, 1, dt)
    sheet1.update_cell(chell, 2, gethostname())
    sheet1.update_cell(chell, 3, version)
    sheet1.update_cell(chell, 4, user_ids)
    layout7 = [[InputText(font=font)],
               [Button("Submit", font=font)]]  # hard beded becuse og a pygui bag
    window2 = Window("Bug Report🪲", layout7, icon=r'C:\popicon.ico')

    while 1:
        event, values = window2.read()
        if event == WIN_CLOSED or event == 'Submit':
            # values = values['-IN-']
            values = str(values)
            sheet1.update_cell(chell, 5, values)
            window2.close()
            break
        window2.close()


# main

font = ("Ubuntu 15 bold")
font2 = ("Consolas 8 bold")
font3 = ("Ubuntu 12 bold")
size = (20, 1)

#time.sleep(60)

scope = ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive']
credentials = {}
# sheets bootloger
gc = service_account_from_dict(credentials)
# google_sh = gc.open("tost")
# sheet1 = google_sh.worksheet("Sheet1")


google_sh = gc.open("pop-popup debugging data")
sheet1 = google_sh.worksheet("Lunch Settings")
FailSafe = sheet1.acell('A2').value
FailSafe = int(FailSafe)
Testing = sheet1.acell('B2').value
Testing = int(Testing)

if Testing == 0:
    bootlogger()



theme('LightBlue')

# layouts for gui
layout0 = [[Text("User:", font=font, size=size)],
           [Button("user1", font=font)], [Button("user2", font=font)],
           [Button("user3", font=font)], [Button("user4", font=font)]]
layout1 = [[Text("Επέλεξε τι τηλεφώνημα ήταν:", font=font)], [Button("Εισερχόμενη 📲", font=font)],
		   [Button("Εξερχόμενη 📞", font=font)],
		   [Button("Εσωτερική 🏢", font=font)],
		   [Checkbox("Δεν απάντησε📵", default=False, key = "-IN-",font=font2)]]
layout2 = [[Text("Επέλεξε τι τύπου ήταν:", font=font)],
           [Button("Delay calls ⏰", font=font)],
           [Button("Missing delivery notes info 🛵", font=font)],
           [Button("Problem with the order 🛒", font=font)],
           [Button("Technical Issue 🛠", font=font)],
           [Button("Other 😶", font=font)]]
layout3 = [[Text("Πληκτρολογήστε γιατί: (Optional)", font=font)], [InputText(font=font)], [Button("Submit", font=font)]]
layout4 = [[Text("Μενού:", font=font)], [Button("Καταχώριση χωρίς F9", font=font), Button("Έξοδος", font=font)]]

# set user
window0 = Window("Select CC agent", layout0, icon=r'C:\popicon.ico')

while 1:
    event, values = window0.read()
    if event == "user1":
        window0.close()
        user_id = 105
        break
    if event == "user2":
        window0.close()
        user_id = 118
        break
    if event == "user3":
        window0.close()
        user_id = 106
        break
    if event == "user4":
        window0.close()
        user_id = 102
        break
    else:
        window0.close()
        user_id = 1955
        break

if Testing == 0:
    loginlogger()

#google_sh = gc.open("CC tracking")
#sheet1 = google_sh.worksheet("PHONETRACKING")

c_a = 0  # call unserde flag
if user_id == 1955:
    flag = 0
else:
    coll = get_from_name(user_id)
    theme(oldtheme())
    name = getname(user_id)
    if Testing == 0:
        google_sh = gc.open("CC tracking")
        sheet1 = google_sh.worksheet("PHONETRACKING")
    elif Testing == 1:
        google_sh = gc.open("tost")
        sheet1 = google_sh.worksheet("Sheet1")
    greedgeexaltedones()

flag = 1



while flag == 1:
    event = keyboard.read_event()
    if event.event_type == keyboard.KEY_DOWN and event.name == 'f9':
        row = get_from_chell()
        call_answerd()
        c_a = 1
    if event.event_type == keyboard.KEY_DOWN and event.name == 'f10':
        if c_a == 1: #benzina
            myfirstmenu()
            c_a = 0
        elif c_a == 0:
            Popup("Δεν έχεις πατήσει F9, πάτα 'end' για να ανοίξεις το μενού", font=font)
    if event.event_type == keyboard.KEY_DOWN and event.name == 'end':
        layout4 = [[Text("Μενού:", font=font)],
                   [Button("Αλλαγή θέματος🖌️", font=font), Button("Καταχώριση χωρίς F9⚠️", font=font)],
                   [Button("Report a Bug🐛", font=font), Button("Έξοδος από την Εφαρμογή", font=font)],
                   [Text("User:", font=font2), Text(name, font=font2), Text(" Version:", font=font2), Text(version, font=font2)]]
        window1 = Window("Μενού", layout4, icon=r'C:\popicon.ico')
        while 1:
            event, values = window1.read()
            if event == "Καταχώριση χωρίς F9⚠️":
                window1.close()
                row = get_from_chell()
                call_not_answerd()
                myfirstmenu()
                break

            if event == "Έξοδος από την Εφαρμογή":
                Popup("Bye bye", name, font=font)
                flag = 0
                break
                window1.close()
            if event == "Report a Bug🐛":
                bugreport()
                if Testing == 0:
                    google_sh = gc.open("CC tracking")
                    sheet1 = google_sh.worksheet("PHONETRACKING")
                elif Testing == 1:
                    google_sh = gc.open("tost")
                    sheet1 = google_sh.worksheet("Sheet1")
            if event == WIN_CLOSED:
                window1.close()
                break
            if event == "Αλλαγή θέματος🖌️":
                window1.close()
                theme1 = change_theme()
                theme(theme1)
                setthemeindb()
                if Testing == 0:
                    google_sh = gc.open("CC tracking")
                    sheet1 = google_sh.worksheet("PHONETRACKING")
                elif Testing == 1:
                    google_sh = gc.open("tost")
                    sheet1 = google_sh.worksheet("Sheet1")
                # theme('DarkPurple5')
                break
