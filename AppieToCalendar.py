from dataclasses import dataclass
from icalendar import Calendar, Event, vCalAddress, vText
from datetime import datetime
from pathlib import Path
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from dateutil.relativedelta import relativedelta

import tkinter as tk

import json

import json
def importData():
    with open("data.json", "r") as jsonFile:
        return json.load(jsonFile)

def saveData(data):
    with open("data.json", "w") as jsonFile:
        json.dump(data, jsonFile)

def syncCalendar(userName, password):
    data = importData()

    options = Options()
    options.add_argument('--headless')
    #options.add_argument('--disable-gpu')
    driver = webdriver.Chrome("/opt/homebrew/bin/chromedriver", chrome_options=options)
    driver.get("https://sam.ahold.com")

    print(driver.current_url)
    
    time.sleep(1)

    passwordInputBox = driver.find_element(By.XPATH, "/html/body/div/main/div[2]/div[2]/div/div/div/div/div/div/div[2]/form/div[1]/div[1]/div/input")
    passwordInputBox.send_keys(userName.get())

    passwordInputBox = driver.find_element(By.XPATH, "/html/body/div/main/div[2]/div[2]/div/div/div/div/div/div/div[2]/form/div[1]/div[2]/div/input")
    passwordInputBox.send_keys(password.get())

    driver.find_element(By.XPATH, "/html/body/div/main/div[2]/div[2]/div/div/div/div/div/div/div[2]/form/div[2]/div/div/input").click()

    driver.get("https://sam.ahold.com/wrkbrn_jct/etm/etmMenu.jsp?locale=nl_NL")

    driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr/td/table[2]/tbody/tr/td/center/table/tbody/tr/td[1]/table/tbody/tr").click()

    newItems = []

    rooster = driver.find_elements(By.CLASS_NAME, "calendarCellRegularFuture")
    for enum, item in enumerate(rooster):
        if enum % 2 == 0:
            print("--==--")
            print(item.text.split("\n")[0], item.text.split("\n")[1])

            parsedTime = item.text.split("\n")[1].split("-")[0].replace("(", "")

            newItem = []

            newItem.append(datetime(datetime.now().year, datetime.now().month, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1])))

            parsedTime = item.text.split("\n")[1].split("-")[1].replace(")", "")

            newItem.append(datetime(datetime.now().year, datetime.now().month, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1])))

            if [str(newItem[0]), str(newItem[1])] not in data["savedTimes"]:
                newItems.append(newItem)


    driver.find_element(By.XPATH, "/html/body/form/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span").click()

    rooster = driver.find_elements(By.CLASS_NAME, "calendarCellRegularFuture")
    for enum, item in enumerate(rooster):
        if enum % 2 == 0:
            print("--==--")
            print(item.text.split("\n")[0], item.text.split("\n")[1])
            
            newItem = []

            parsedTime = item.text.split("\n")[1].split("-")[0].replace("(", "")

            newItem.append(datetime(datetime.now().year, datetime.now().month, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1]))+ relativedelta(months=1))

            parsedTime = item.text.split("\n")[1].split("-")[1].replace(")", "")

            newItem.append(datetime(datetime.now().year, datetime.now().month, int(item.text.split("\n")[0]), int(parsedTime.split(":")[0]), int(parsedTime.split(":")[1]))+ relativedelta(months=1))

            if [str(newItem[0]), str(newItem[1])] not in data["savedTimes"]:
                newItems.append(newItem)
    
    driver.quit()

    print(newItems)

    cal = Calendar()
    cal.add('prodid', '-//My calendar product//example.com//')
    cal.add('version', '2.0')

    for i in newItems:
        event = Event()

        #add properties to event
        event.add('summary', 'Werken')
        #event.add('description', 'Je moet werken')
        event.add
        event.add('dtstart', i[0])
        event.add('dtend', i[1])
        cal.add_component(event)

    f = open(os.path.join("./", 'example.ics'), 'wb')
    f.write(cal.to_ical())
    f.close()

    for item in newItems:
        data["savedTimes"].append([str(item[0]), str(item[1])])
    
    tk.Label(text="Ik heb {} werktijden gevonden!".format(len(newItems))).grid(row = 4, column = 0, columnspan = 2)

    saveData(data)
    

screen = tk.Tk()

tk.Label(text="User ID").grid(row = 1, column = 0)
userNameEntry = tk.Entry()
userNameEntry.grid(row = 1, column = 1)

tk.Label(text="Password").grid(row = 2, column = 0)
passwordEntry = tk.Entry()
passwordEntry.grid(row = 2, column = 1)

tk.Button(text="Exporteer appie rooster", command=lambda:syncCalendar(userNameEntry, passwordEntry)).grid(row = 3, column = 0, columnspan= 2)

screen.mainloop()
screen.quit()