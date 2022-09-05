#Time/calendar imports
from re import I
import string
from tokenize import String
from icalendar import Calendar, Event
from datetime import date, datetime

#Webscraping imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from dateutil.relativedelta import relativedelta

#GUI imports
import tkinter as tk

#Data imports
import json

def importData():
    with open("data.json", "r") as jsonFile:
        return json.load(jsonFile)

def saveData(data):
    with open("data.json", "w") as jsonFile:
        json.dump(data, jsonFile)

def exportToIcalendar(items):
    #Initiates Calendar file (.iso)
    cal = Calendar()
    cal.add('prodid', '-//My calendar product//example.com//')
    cal.add('version', '2.0')

    #Makes a calendar event for each planned workday
    print(items)
    for i in items:
        event = Event()
        event.add('summary', 'Werken bij AH')
        event.add('dtstart', i[0])
        event.add('dtend', i[1])
        cal.add_component(event)

    #Exports Icalendar file
    f = open("./example.ics", "wb")
    f.write(cal.to_ical())
    f.close()

def syncCalendar(userName, password):
    data = importData()
    
    # >>> Put chromedriver path here <<<
    chromeDriverPath = "/opt/homebrew/bin/chromedriver"

    #Initiate driver
    options = Options()
    options.add_argument('--headless')
    driver = webdriver.Chrome(chromeDriverPath, chrome_options=options)
    driver.get("https://sam.ahold.com")
    
    #Don't question this, without it it doesn't work sometimes
    time.sleep(1)

    #Goes to the right page (Login > Main page > schedule page)
    passwordInputBox = driver.find_element(By.XPATH, "/html/body/div/main/div[2]/div[2]/div/div/div/div/div/div/div[2]/form/div[1]/div[1]/div/input")
    passwordInputBox.send_keys(userName.get())
    passwordInputBox = driver.find_element(By.XPATH, "/html/body/div/main/div[2]/div[2]/div/div/div/div/div/div/div[2]/form/div[1]/div[2]/div/input")
    passwordInputBox.send_keys(password.get())
    driver.find_element(By.XPATH, "/html/body/div/main/div[2]/div[2]/div/div/div/div/div/div/div[2]/form/div[2]/div/div/input").click()
    driver.get("https://sam.ahold.com/wrkbrn_jct/etm/etmMenu.jsp?locale=nl_NL")
    driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td/table[2]/tbody/tr/td/center/table/tbody/tr/td[1]/table/tbody/tr").click()

    #Scans for planned work times and adds them to list
    newItems = []
    schedule = driver.find_elements(By.CLASS_NAME, "calendarCellRegularFuture")
    for enum, item in enumerate(schedule):
        #Times are double in html, this makes sure only halve are used
        if enum % 2 == 0:
            parsedTime = item.text.split("\n")[1].split("-")[0].replace("(", "")
            newItem = []
            newItem.append(datetime(datetime.now().year, datetime.now().month, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1])))
            parsedTime = item.text.split("\n")[1].split("-")[1].replace(")", "")
            newItem.append(datetime(datetime.now().year, datetime.now().month, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1])))
            if [str(newItem[0]), str(newItem[1])] not in data["savedTimes"]:
                newItems.append(newItem)

    #Pushes button to go to next month
    driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span").click()

    #Scrapes data for next month (just in case)
    schedule = driver.find_elements(By.CLASS_NAME, "calendarCellRegularFuture")
    for enum, item in enumerate(schedule):
        #Times are double in html, this makes sure only halve are used
        if enum % 2 == 0:
            newItem = []
            parsedTime = item.text.split("\n")[1].split("-")[0].replace("(", "")
            newItem.append(datetime(datetime.now().year, datetime.now().month, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1]))+ relativedelta(months=1))
            parsedTime = item.text.split("\n")[1].split("-")[1].replace(")", "")
            newItem.append(datetime(datetime.now().year, datetime.now().month, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1]))+ relativedelta(months=1))
            if [str(newItem[0]), str(newItem[1])] not in data["savedTimes"]:
                newItems.append(newItem)
    
    #Quits drives, so it doesn't just sit doing nothing
    driver.quit()
    
    exportToIcalendar(newItems)

    #Adds saved timestamps to json to prevent duplicate exports and enable work statistics
    for item in newItems:
        data["savedTimes"].append([str(item[0]), str(item[1])])
    saveData(data)
    
    tk.Label(text="Ik heb {} werktijden gevonden!".format(len(newItems))).grid(row = 4, column = 0, columnspan = 2)

def clearScreen():
    for widget in screen.winfo_children():
        widget.destroy()

def syncCalendarGUI():
    clearScreen()

    tk.Button(text="Terug", command=mainScreen).grid(row = 0, column = 0)
    tk.Label(text="Automatisch rooster ophalen").grid(row = 0, column = 1)

    tk.Label(text="User ID").grid(row = 1, column = 0)
    userNameEntry = tk.Entry()
    userNameEntry.grid(row = 1, column = 1)

    tk.Label(text="Password").grid(row = 2, column = 0)
    passwordEntry = tk.Entry()
    passwordEntry.grid(row = 2, column = 1)

    #Passes userName and password for use during the scraping
    tk.Button(text="Exporteer appie rooster", command=lambda:syncCalendar(userNameEntry, passwordEntry)).grid(row = 3, column = 0, columnspan= 2)

def popup(message):
    popupScreen = tk.Toplevel()
    
    popupScreen.title("Error!")
    tk.Label(popupScreen, text=message).pack()
    tk.Button(popupScreen, text="Begrepen!", command=lambda:popupScreen.destroy()).pack()

    popupScreen.mainloop()

def addNewWorkTime(datumEntry, beginTimeEntry, endTimeEntry, newEventCounter, changeCounterLabel, data, newData):
    try:
        datum = datumEntry.get()
        beginTime = beginTimeEntry.get()
        endTime = endTimeEntry.get()

        if len(datum.split("-")) != 3:
            popup("Datum is niet correct opgeschreven \n voorbeeld van correcte datum: 2022-5-23")
        elif len(beginTime.split(":")) != 2 or len(endTime.split(":")) != 2:
            popup("Tijd is niet correct opgeschreven \n voorbeeld van correcte tijd notatie: 14:23")
        else:
            year = int(datum.split("-")[0])
            month = int(datum.split("-")[1])
            day = int(datum.split("-")[2])
            beginDateTime = datetime(year, month, day, int(beginTime.split(":")[0]), int(beginTime.split(":")[1]), 0)
            endDateTime = datetime(year, month, day, int(endTime.split(":")[0]), int(endTime.split(":")[1]), 0)
            data["savedTimes"].append([str(beginDateTime), str(endDateTime)])
            newData.append([beginDateTime, endDateTime])
            newEventCounter += 1
            changeCounterLabel.config(text = str(newEventCounter) + " werktijd(en) toegevoegd")
            saveData(data)
    except Exception as e:
        popup("Je hebt iets fout gedaan:\n python error " + str(e))
        
def manuallyAddWorkTime():
    clearScreen()

    data = importData()

    tk.Button(text="Terug", command=mainScreen).grid(row = 0, column = 0)
    tk.Label(text="Zelf werktijd(en) invullen").grid(row = 0, column = 1)

    tk.Label(text="datum (jaar-maand-dag)").grid(row = 1, column = 0)
    datumEntry = tk.Entry()
    datumEntry.grid(row = 1, column = 1)

    tk.Label(text="begin tijd (uur:minuut)").grid(row = 2, column = 0)
    beginTimeEntry = tk.Entry()
    beginTimeEntry.grid(row = 2, column = 1)

    tk.Label(text="eind tijd (uur:minuut)").grid(row = 3, column = 0)
    endTimeEntry = tk.Entry()
    endTimeEntry.grid(row = 3, column = 1)

    changeCounterLabel = tk.Label(text="0 werktijden toegevoegd")
    changeCounterLabel.grid(row = 5, column = 0, columnspan= 2)

    newEventCounter = 0
    newData = []
    tk.Button(text="Voeg toe", command=lambda:addNewWorkTime(datumEntry, beginTimeEntry, endTimeEntry, newEventCounter, changeCounterLabel, data, newData)).grid(row = 4, column = 0, columnspan=2)

    tk.Button(text="Exporteer nieuwe tijden naar agenda", command=lambda:exportToIcalendar(newData)).grid(row = 6, column = 0, columnspan=2)

def mainScreen():
    #Initiate main menu GUI
    clearScreen()
    tk.Button(text="Automatisch rooster ophalen", command=lambda:syncCalendarGUI()).pack()
    tk.Button(text="Zelf werktijd(en) invullen", command=lambda:manuallyAddWorkTime()).pack()

screen = tk.Tk()
screen.title("Appie to Calendar V0.2")
screen.resizable(False, False) 
mainScreen()
screen.mainloop()
screen.quit()