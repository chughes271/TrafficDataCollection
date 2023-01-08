#IMPORT SELENIUM LIBRARIES IN ORDER TO SCRAPE DATA FROM THE INTERNET
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

#IMPORT TIME LIBRARIES IN ORDER TO RECORD AND EXPORT TIME OF RUN
from datetime import date
from datetime import datetime
import time

#IMPORT PANDAS LIBRARY IN ORDER TO MANIPULATE THE RECORDED DATA AND EXPORT TO EXCEL PROPERLY
import pandas as pd


#CREATE LOCATIONS USED AND PATH OF EXCEL SHEET AS GLOBAL VARIABLES
home = "2814 Guillot st, Dallas TX"
work = "Innovation First International, FM 1570, Greenville, TX"
path = 'E:\Code\TrafficPredictor\LongTermTrafficData.xlsx'
morningtimes = ["0701", "0731", "0801", "0831", "0901", "0931"]
afternoontimes = ["1601", "1631", "1701", "1731", "1801", "1831"]

def getGasPrice():
##---GAS PRICE FUNCTION - GOES TO AAA WEBSITE AND RECORDS THE CURRENT GAS PRICE IN TEXAS---##

    #Go to Texas gas price page
    driver.get('https://gasprices.aaa.com/?state=TX')

    #Locate state wide gas price on current page
    gasPosition = driver.find_element(By.XPATH, '//*[@id="maincontent"]/div[3]/div/div[1]/div[2]/p[2]')
    gasPrice = gasPosition.get_property("innerText")

    #Extracts just numerical value from string and changes type to number
    gasPrice = gasPrice.removeprefix("$")
    gasPrice = gasPrice.removesuffix(" ")
    gasPrice = round(float(gasPrice),2)

    return(gasPrice)

def getTimeDayDate(type):
##---TIME FUNCTION TO RECORD ANY DATA ON DAY,TIME, DATE---##

    #0= current time
    if type == 0:
        now = datetime.now()
        output = now.strftime("%H%M")

    # 1= day - returns number (0 through 6) 0=monday ... 6=sunday
    elif type == 1:
        today = date.today()
        output = today.weekday()

    # 2= date
    elif type == 2:
        today = date.today()
        output = today.strftime("%m%d%y")

    # 3= day, day and time together
    elif type == 3:
        now = datetime.now()
        dateAndTime = now.strftime("%m/%d/%Y, %H:%M:%S")
        today = date.today()
        output = str(today.weekday()) +", "+ dateAndTime

    return(output)

def getWeekdayName():
##---CHANGE FUNCTION FROM VALUE (0-6) TO WEEKDAYNAME (MONDAY-SUNDAY)---##
    weekdayNum = getTimeDayDate(1)

    return{
        0: "Monday",
        1: "Tuesday",
        2: "Wednesday",
        3: "Thursday",
        4: "Friday",
        5: "Saturday",
        6: "Sunday"
    }[weekdayNum]


def locationInput():
##---MAIN TRAFFIC SCRAPING CODE---##
    #open Google maps
    driver.get("https://www.google.com/maps/dir/VEX+Robotics,+Inc.,+6725+FM+1570,+Greenville,+TX+75402/Dallas,+TX/@32.9409533,-96.5725588,10.72z/data=!4m13!4m12!1m5!1m1!1s0x864be67e938bd9ab:0x1289986db0c5ae01!2m2!1d-96.149662!2d33.0639453!1m5!1m1!1s0x864c19f77b45974b:0xb9ec9ba4f647678f!2m2!1d-96.7969879!2d32.7766642")
    time.sleep(5)

    #get current time
    timeofRun = getTimeDayDate(0)

    #Check the current hour value and move forward accordingly
    #IF MORNING (7-9)am - Input home as starting point and job as destination
    #IF AFTERNOON (16-18)pm - Input job as starting point and home as destination

    ##---MORNING RUN---## (Home going to work)
    if timeofRun[0:2] >= "07" and timeofRun[0:2] <= "09":

        #LOOK FOR STARTING POSITION INPUT BOX
        startingPointBox = driver.find_element(By.XPATH, '//*[@id="sb_ifc50"]/input')
        startingPointBox.clear()

        #INPUT HOME INTO STARTING POSITION
        startingPointBox.send_keys(home)
        startingPointBox.send_keys(Keys.RETURN)


        #LOOK FOR DESTINATION INPUT BOX
        startingPointBox = driver.find_element(By.XPATH, '//*[@id="sb_ifc51"]/input')
        startingPointBox.clear()

        #INPUT JOB INTO DESTINATION POSITION
        startingPointBox.send_keys(work)
        startingPointBox.send_keys(Keys.RETURN)

        time.sleep(10)

    ##---AFTERNOON RUN---## (Work going home)
    elif timeofRun[0:2] >= "16" and timeofRun[0:2] <= "18":

        #LOOK FOR STARTING DESTINATION INPUT BOX
        startingPointBox = driver.find_element(By.XPATH, '//*[@id="sb_ifc51"]/input')
        startingPointBox.clear()

        #INPUT HOME INTO DESTINATION
        startingPointBox.send_keys(home)
        startingPointBox.send_keys(Keys.RETURN)

        #LOOK FOR STARTING POINT INPUT BOX
        startingPointBox = driver.find_element(By.XPATH, '//*[@id="sb_ifc50"]/input')
        startingPointBox.clear()

        #INPUT JOB INTO STARTING POINT
        startingPointBox.send_keys(work)
        startingPointBox.send_keys(Keys.RETURN)

        time.sleep(10)
    else:
        #ERROR CHECK - PRINT TIME WAS NOT IN THE IF STATEMENT BOUNDS
        #THIS WILL CONTINUE TO RUN (WORK TO HOME)
        print("Time niether 7-9am or 4-6pm")

def getDuration():
##---FUNCTION TO SCRAPE DURATION VALUE---##
    #FIND DURATION VALUE ON CURRENT PAGE
    durationPath = driver.find_element(By.XPATH, '//*[@id="section-directions-trip-0"]/div[1]/div[1]/div[1]/div[1]/span[1]')
    durationValue = str(durationPath.get_property('innerHTML'))

    #INTAKE AND PROCESS DATA TO FIND VALUE IN MINUTES
    if len(durationValue) > 6: #IF LENGTH OF VALUE >6 THEN VALUE IS LONGER THAN 1HR (EX: 1hr 34min)
        #EXTRACT NUMERICAL VALUES
        durationValue = durationValue.removesuffix(" min")
        durationValue = durationValue.replace(" hr ","",1)

        #TURN VALUE INTO MINUTE VALUE (EX 1HR 12MIN CHANGES TO 72MIN)
        durationValue = float(durationValue[0]) * 60 + float(durationValue[1:len(durationValue)+1])

    elif len(durationValue) == 4: #IF LENGTH OF VALUE IS =4 THEN VALUE IS HOUR INTEGER (EX: 2 hr OR 9 hr)
        #EXTRACT NUMERICAL VALUES
        durationValue = durationValue.removesuffix(" hr")
        #TURN VALUE INTO MINUTE VALUE
        durationValue = float(durationValue) * 60

    else: #ELSE MEANS DATA IS LESS THAN AN HOUR (EX: 45 MIN)
        #EXTRACT VALUE AND TURN TYPE INTO NUMBER
        durationValue = durationValue.removesuffix(" min")
        durationValue = float(durationValue)
    #print(durationValue)
    return(durationValue)

def getDistance():
##---FUNCTION TO GET DISTANCE VALUE---##
    #LOCATE DISTANCE VALUE IN CURRENT PAGE
    distancePath = driver.find_element(By.XPATH, '//*[@id="section-directions-trip-0"]/div[1]/div[1]/div[1]/div[2]/div')
    distanceValue = (distancePath.get_property('innerText'))
    #EXTRACT NUMERICAL VALUE
    distanceValue = distanceValue.removesuffix(" miles")
    distanceValue = float(distanceValue)

    return(distanceValue)

def gasDataCollection(gasPrice,distanceValue):
##---FUNCTION TO PROCESS GAS DATA FOR TWO CARS (2021 JEEP GRAND CHEROKEE AND 2022 TOYOTA RAV4)---##
    #First number is JEEP
    #Second number is Rav4
    mpgCombined = [22.5,29.0]
    mpgCity = [19.0,27.0]
    mpgHw = [26.0,33.0]

    #Calculate price of drive by Gas PER DAY for the cars
    combinedGasPrice = [gasPrice * distanceValue / i for i in mpgCombined]
    cityGasPrice = [gasPrice * distanceValue / i for i in mpgCity]
    hwGasPrice = [gasPrice * distanceValue / i for i in mpgHw]

    return(combinedGasPrice,cityGasPrice,hwGasPrice)


def packageExport(distance,duration,gasPrice,combinedGasPrice,cityGasPrice,hwGasPrice):
##---FUNCTION TO PACKAGE ALL DATA PROCESSED IN THE RUN READY FOR EXCEL EXPORT---##

    #ESTABLISH RUN TIMES AND TODAY'S TIME DATA
    timeVector = getTimeDayDate(3)
    timeofRun = getTimeDayDate(0)

    #CREATE PACKAGE TO EXPORT
    dataList = {'Time': duration, 'Distance':distance, 'Combined Gas Price Per Day Jeep':combinedGasPrice[0], 'Optimistic Gas Price Per Day Jeep':hwGasPrice[0],
    'Pessimistic Gas Price Per Day Jeep':cityGasPrice[0], 'Combined Gas Price Per Day RAV':combinedGasPrice[1],'Optimistic Gas Price Per Day RAV':hwGasPrice[1],
    'Pessimistic Gas Price Per Day RAV':cityGasPrice[1],'Gas Price':gasPrice,'Time Units':timeVector}

    #TURN PACKAGE INTO DATAFRAME IN PANDAS
    df = pd.DataFrame(dataList,index=[0])
    col = len(df.axes[1])

    #ASK WHAT TIME IT IS BY MATCHING RUN TIME TO PREDETERMINED TIMES TO DECIDE EXCEL LOCATION TO EXPORT TO
    for x in range(0,len(morningtimes)):
        if morningtimes[x] == timeofRun:
            startcol = col*x
            break
        elif afternoontimes[x] == timeofRun:
            startcol = col * (x+len(morningtimes)-1)
            break
        else: #ELSE PUT INTO AN UNUSED COLUMN SO DATA CAN STILL BE SEEN WITHOUT INTERFERING WITH OTHER DATA
            startcol = col * 12

    return(df,startcol)

def readExcel():
##---EXCEL READING FUNCTION TO DETERMINE ROW LOCATION (IF ITS A NEW WEEK)---##
    #GETTIME DATA
    day = getTimeDayDate(1)
    timeNow = getTimeDayDate(0)

    #READ THE ROW LENGTH OF THE "MONDAY" EXCEL TAB
    data = pd.read_excel(path, 'Monday')
    nthRow = (len(data))

    #START A NEW ROW IF TIME IS 7AM ON MONDAY (THE FIRST RUN OF THE DAY AND WEEK)
    if timeNow == '0701' and day == 0:
        startRow = nthRow + 2

    #KEEP THE SAME ROW IF THE RUN IS NOT 7AM ON MONDAY
    else:
        startRow = round(nthRow - 1)

    return(startRow)


def exportToExcel(df,startRow,startCol,weekDayName):
##---FUNCTION TO EXPORT PACKAGE TO EXCEL---##
    #EXPORT PACKGE AT "STARTROW" AND "STARTCOL" IN THE SHEET THAT MATCHES THE DAY OF THE WEEK
    with pd.ExcelWriter(path, mode='a', if_sheet_exists='overlay') as book:
        df.to_excel(book, sheet_name=weekDayName, startrow=startRow, startcol=startCol, index=False)

##---CALL ALL FUNCTIONS---##
driver = webdriver.Chrome()
weekDay = getWeekdayName()
startRow = readExcel()
gasPrice = getGasPrice()
locationInput()
duration = getDuration()
distance = getDistance()
[combinedGasPrice,cityGasPrice,hwGasPrice] = gasDataCollection(gasPrice,distance)
package,startCol = packageExport(distance,duration,gasPrice,combinedGasPrice,cityGasPrice,hwGasPrice)
exportToExcel(package,startRow,startCol,weekDay)
driver.close()