#Time/calendar imports
from dataclasses import dataclass
import enum
from itertools import count
from re import I
import string
from tokenize import String
from turtle import clear
from icalendar import Calendar, Event
from datetime import date, datetime

#Webscraping imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from dateutil.relativedelta import relativedelta
import selenium.common.exceptions

#GUI imports
import tkinter as tk
from tkinter import ttk 
#Data imports
import json

class ScrollableFrame(ttk.Frame):
    global background_color
    def __init__(self, container, borderwidth_in, frame_size, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, width = frame_size[0], height = frame_size[1], borderwidth = borderwidth_in)
        scrollbar_y = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        scrollbar_x = tk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.scrollable_frame = tk.Frame(self.canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        self.canvas.grid(sticky="e", row = 0, column = 0)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='we')
    
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def move_down(self):
        self.canvas.yview_moveto(1)

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

def syncCalendar(userName, password, statusLabel, monthsToGoBack, deleteSavedTimes):

    statusLabel.config(text="Bezig (duurt maximaal 1 minuut)...")
    screen.update()

    if deleteSavedTimes.get() == 0:
        data = importData()
    else:
        data = {"savedTimes": []}

    print("blup")

    # >>> Put chromedriver path here <<<
    chromeDriverPath = "/opt/homebrew/bin/chromedriver"

    #Initiate driver
    options = Options()
    options.add_argument('--headless')
    driver = webdriver.Chrome(chromeDriverPath, chrome_options=options)
    try:
        driver.get("https://sam.ahold.com")
    except selenium.common.exceptions.WebDriverException:
        popup("Geen internet connectie!")

    print("Driver innitiated")
    #Goes to the right page (Login > Main page > schedule page)
    timeOutCounter = 0
    while len(driver.find_elements(By.XPATH, "/html/body/div/main/div[2]/div[2]/div/div/div/div/div/div/div[2]/form/div[1]/div[1]/div/input")) == 0 and timeOutCounter <= 50:
        time.sleep(0.2)
        timeOutCounter += 1
        if timeOutCounter > 50:
            popup("Timeout terwijl browser werd opgestart.")
    passwordInputBox = driver.find_element(By.XPATH, "/html/body/div/main/div[2]/div[2]/div/div/div/div/div/div/div[2]/form/div[1]/div[1]/div/input")
    passwordInputBox.send_keys(userName.get())
    passwordInputBox = driver.find_element(By.XPATH, "/html/body/div/main/div[2]/div[2]/div/div/div/div/div/div/div[2]/form/div[1]/div[2]/div/input")
    passwordInputBox.send_keys(password.get())
    try:
        driver.find_element(By.XPATH, "/html/body/div/main/div[2]/div[2]/div/div/div/div/div/div/div[2]/form/div[2]/div/div/input").click()
        driver.get("https://sam.ahold.com/wrkbrn_jct/etm/etmMenu.jsp?locale=nl_NL")
        if len(driver.find_elements(By.XPATH, "/html/body/form/table/tbody/tr/td/table[2]/tbody/tr/td/center/table/tbody/tr/td[1]/table/tbody/tr")) == 0:
            popup("Incorrecte gebruikersnaam/wachtwoord ingevuld!")
        driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td/table[2]/tbody/tr/td/center/table/tbody/tr/td[1]/table/tbody/tr").click()
    except selenium.common.exceptions.UnexpectedAlertPresentException:
        popup("Geen gebruikersnaam/wachtwoord ingevuld!")
    print("Username succesfully entered")

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
    print("Scan completed")
    
    oldTimes = []

    try:
        if int(monthsToGoBack.get()) >= 0:
            for i in range(int(monthsToGoBack.get()) + 1):
                schedule = driver.find_elements(By.CLASS_NAME, "calendarCellRegularPast")
                for enum, item in enumerate(schedule):
                    if enum % 2 == 0:
                        parsedTime = item.text.split("\n")[1].split("-")[0].replace("(", "")
                        newItem = []
                        print(int(item.text.split("\n")[0]))
                        newItem.append(datetime(datetime.now().year, 1, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1])) + relativedelta(months = datetime.now().month - i - 1))
                        parsedTime = item.text.split("\n")[1].split("-")[1].replace(")", "")
                        newItem.append(datetime(datetime.now().year, 1, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1])) + relativedelta(months = datetime.now().month - i - 1))
                        if [str(newItem[0]), str(newItem[1])] not in data["savedTimes"]:
                            oldTimes.append(newItem)
                driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td[1]/table/tbody/tr").click()
            
            for i in range(int(monthsToGoBack.get()) + 1):
                driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span").click()
        else:
            popup(monthsToGoBack.get() + " is niet een correct getal.")
    except ValueError:
        popup(monthsToGoBack.get() + " is niet een correct getal.")

    #Pushes button to go to next month
    driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span").click()

    #Scrapes data for next month (just in case)
    schedule = driver.find_elements(By.CLASS_NAME, "calendarCellRegularFuture")
    for enum, item in enumerate(schedule):
        #Times are double in html, this makes sure only halve are used
        if enum % 2 == 0:
            newItem = []
            parsedTime = item.text.split("\n")[1].split("-")[0].replace("(", "")
            newItem.append(datetime(datetime.now().year, datetime.now().month, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1])) + relativedelta(months = 1))
            parsedTime = item.text.split("\n")[1].split("-")[1].replace(")", "")
            newItem.append(datetime(datetime.now().year, datetime.now().month, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1])) + relativedelta(months = 1))
            if [str(newItem[0]), str(newItem[1])] not in data["savedTimes"]:
                newItems.append(newItem)
    
    #Quits drives, so it doesn't just sit doing nothing
    driver.quit()
    
    exportToIcalendar(newItems)

    #Adds saved timestamps to json to prevent duplicate exports and enable work statistics
    for item in newItems:
        data["savedTimes"].append([str(item[0]), str(item[1])])

    for item in oldTimes:
        data["savedTimes"].append([str(item[0]), str(item[1])])
    saveData(data)

    statusLabel.config(text="Heb {} nieuwe en {} oude werktijden gevonden".format(len(newItems), len(oldTimes)))

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
    passwordEntry = tk.Entry(show="*")
    passwordEntry.grid(row = 2, column = 1)

    tk.Label(text="Hoeveelheid verleden maanden ophalen").grid(row = 3, column = 0)
    monthsToGoBack = tk.Entry()
    monthsToGoBack.grid(row = 3, column=1)
    monthsToGoBack.insert(0, "5")

    deleteSavedTimes = tk.IntVar()
    tk.Checkbutton(text="Verwijder opgeslagen werk tijden", variable=deleteSavedTimes).grid(row = 4, column = 0, columnspan=2)

    
    statusLabel = tk.Label(text="")
    statusLabel.grid(row = 6, column = 0, columnspan = 2)

    #Passes userName and password for use during the scraping
    tk.Button(text="Exporteer appie rooster", command=lambda:syncCalendar(userNameEntry, passwordEntry, statusLabel, monthsToGoBack, deleteSavedTimes)).grid(row = 5, column = 0, columnspan= 2)

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

def Statistics():
    clearScreen()

    data = importData()

    tk.Button(text="Terug", command=mainScreen).grid(row = 0, column = 0)
    tk.Label(text="Statistieken").grid(row = 0, column = 1)

    totalTime = 0
    payoutTime = 0
    timeData = {}
    for item in data["savedTimes"]:
        dateTimeDifference = str(datetime(2000, 1, 1, int(item[1].split(" ")[1].split(":")[0]), int(item[1].split(" ")[1].split(":")[1]), 0) - datetime(2000, 1, 1, int(item[0].split(" ")[1].split(":")[0]), int(item[0].split(" ")[1].split(":")[1]), 0))
        minuteDifference = int(dateTimeDifference.split(":")[0]) * 60 + int(dateTimeDifference.split(":")[1])
        if item[0].split("-")[0] not in timeData.keys() or item[0].split("-")[1] not in timeData[item[0].split("-")[0]].keys():
            if item[0].split("-")[0] not in timeData.keys():
                timeData[item[0].split("-")[0]] = {}
            timeData[item[0].split("-")[0]][item[0].split("-")[1]] = minuteDifference / 60
        else:
             timeData[item[0].split("-")[0]][item[0].split("-")[1]] += minuteDifference / 60
        totalTime += minuteDifference / 60
        #print(int(item[0].split("-")[2][0:1]))
        print(datetime(int(item[0].split("-")[0]), int(item[0].split("-")[1]), int(item[0].split("-")[2][0:2])))
        if datetime(int(item[0].split("-")[0]), int(item[0].split("-")[1]), int(item[0].split("-")[2][0:2])).weekday() == 6:
            payoutTime += minuteDifference / 60.0 * 1.5
        else:
            payoutTime += minuteDifference / 60.0
    print(timeData)

    tk.Label(text="Totaal gewerkte uren: " + str(totalTime)).grid(row = 1, column = 0, columnspan = 2)
    tk.Label(text="Totaal uitbetaalde uren: " + str(payoutTime)).grid(row = 2, column = 0, columnspan = 2)
    tk.Label(text="Totaal verdient (ongeveer): " + str(payoutTime * 6.84)).grid(row = 3, column = 0, columnspan = 2)

    tk.Button(text="Uur gewerkt per maand", command=lambda:showWorkGraph(timeData)).grid(row = 4, column = 0, columnspan = 2)
    tk.Button(text="Uur gewerkt per uitbetaal periode").grid(row = 5, column = 0, columnspan = 2)
    tk.Button(text="Change payout settings").grid(row = 6, column = 0, columnspan=2)

def showWorkGraph(timeData):
    import matplotlib.pyplot as plt

    yearList = list(timeData.keys())
    yearList.sort()

    yAxis = []
    xAxis = []
    for year in yearList:
        newMonthList = list(timeData[year].keys())
        newMonthList.sort()

        for month in newMonthList:
            yAxis.append(timeData[year][month])
            xAxis.append(year + "-" + month)
    
    plt.title("Hoeveelheid gewerkt bij de appie")
    plt.xlabel("Tijd")
    plt.ylabel("Aantal gewerkte uren")
    plt.plot(xAxis,yAxis, color='maroon', marker='o')
    plt.show()

def saveEdittedTimes(data, dateList, sortedDateList, beginTimeEntryList, endTimeEntryList):
    for counter, item in enumerate(sortedDateList):
        beginTimeList = [int(beginTimeEntryList[counter].get().split(":")[0]), int(beginTimeEntryList[counter].get().split(":")[1])]
        endTimeList = [int(endTimeEntryList[counter].get().split(":")[0]), int(endTimeEntryList[counter].get().split(":")[1])]
        data["savedTimes"][dateList.index(item)] =[str(datetime(int(item.split("-")[0]), int(item.split("-")[1]), int(item.split("-")[2]), beginTimeList[0], beginTimeList[1])), str(datetime(int(item.split("-")[0]), int(item.split("-")[1]), int(item.split("-")[2]), endTimeList[0], endTimeList[1]))]
    
    saveData(data)

def exportOldTimes(data, dateList, sortedDateList, syncToggleList):
    exportDates = []
    for counter, item in enumerate(sortedDateList):
        if syncToggleList[counter].get() == 1:
            print(data["savedTimes"][dateList.index(item)])
            newDatetime = []
            newDate = data["savedTimes"][dateList.index(item)][0].split(" ")[0]
            beginTime = data["savedTimes"][dateList.index(item)][0].split(" ")[1]
            endTime = data["savedTimes"][dateList.index(item)][1].split(" ")[1]
            newDatetime.append(datetime(int(newDate.split("-")[0]), int(newDate.split("-")[1]), int(newDate.split("-")[2]), int(beginTime.split(":")[0]), int(beginTime.split(":")[1])))
            newDatetime.append(datetime(int(newDate.split("-")[0]), int(newDate.split("-")[1]), int(newDate.split("-")[2]), int(endTime.split(":")[0]), int(endTime.split(":")[1])))
            exportDates.append(newDatetime)
    exportToIcalendar(exportDates)

def editWorkTimes():
    clearScreen()

    tk.Button(text="Terug", command=mainScreen).grid(row = 0, column = 0)
    tk.Label(text="Zelf werktijd(en) invullen").grid(row = 0, column = 1)

    workTimeFrame = ScrollableFrame(screen, 0, [500, 750])
    workTimeFrame.grid(row = 1, column = 0, columnspan=2)
    data = importData()
    dateList = []
    beginTimeList = []
    endTimeList = []
    for item in data["savedTimes"]:
        dateList.append(item[0].split(" ")[0])
        beginTimeList.append(item[0].split(" ")[1])
        print(item)
        endTimeList.append(item[1].split(" ")[1])
    
    print(beginTimeList)
    print(endTimeList)

    sortedDateList = dateList.copy()
    sortedDateList.sort()
    sortedDateList.reverse()
    
    beginTimeEntryList = []
    endTimeEntryList = []
    syncToggleList = []

    for counter, item in enumerate(sortedDateList):
        tk.Label(workTimeFrame.scrollable_frame, text = item).grid(row = counter, column = 0)
        beginTimeEntryList.append(tk.Entry(workTimeFrame.scrollable_frame))
        beginTimeEntryList[-1].grid(row = counter, column = 1)
        beginTimeEntryList[-1].insert(0, beginTimeList[dateList.index(item)])
        endTimeEntryList.append(tk.Entry(workTimeFrame.scrollable_frame))
        endTimeEntryList[-1].grid(row = counter, column = 2)
        endTimeEntryList[-1].insert(0, endTimeList[dateList.index(item)])
        syncToggleList.append(tk.IntVar())
        tk.Checkbutton(workTimeFrame.scrollable_frame, variable=syncToggleList[-1]).grid(row = counter, column = 3)
    
    tk.Button(text="Save", command=lambda:saveEdittedTimes(data, dateList, sortedDateList, beginTimeEntryList, endTimeEntryList)).grid(row = 2, column = 0)
    tk.Button(text = "Exporteer gekozen tijden naar calender", command = lambda:exportOldTimes(data, dateList, sortedDateList, syncToggleList)).grid(row = 2, column = 1)

def mainScreen():
    #Initiate main menu GUI
    clearScreen()
    tk.Button(text="Automatisch rooster ophalen", command=lambda:syncCalendarGUI()).pack()
    tk.Button(text="Zelf werktijd(en) invullen", command=lambda:manuallyAddWorkTime()).pack()
    tk.Button(text="Wijzig werktijden", command=lambda:editWorkTimes()).pack()
    tk.Button(text="Statistieken", command=lambda:Statistics()).pack()

screen = tk.Tk()
screen.title("Appie to Calendar V0.5")
screen.resizable(False, False) 
mainScreen()
screen.mainloop()
screen.quit()