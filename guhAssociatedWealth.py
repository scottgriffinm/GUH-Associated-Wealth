'''
SCRIPT GAMEPLAN: RUN WEEKLY
1. Reads if investor is active... If yes, [DONE]

2. Fetches data for each investor on
 - current investment dollar value [DONE]
 - investor's performance (weekly/all time) [DONE]
 - GAW's performance (weekly/all time) [DONE]
 
3. Send emails to investors including
 - current investment dollar value [DONE]
 - investor's performance (weekly/all time) [DONE]
 - GAW's performance (weekly/all time) [DONE]

'''

import sys
import time
import smtplib, ssl
import requests
import random
from string import digits
import numpy as np
import pandas as pd
import yfinance as yf
import xlrd
import xlwt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


print('[Waiting...]')

t = [1.2,3.5,5.3,6.6,2.3]
time.sleep(random.choice(t)) # Delay at beginning of program to dodge bot detection

print('[Running script...]')

# This fixes the xlrd/xlwt module
xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True


# create empty dataframe for investor data
dfINV = pd.DataFrame({
    'Name':[],
    'Email':[],
    'SMS Email':[],
    'IIDate':[],
    'IIAmount':[],
    'P1 Allocation':[],
    'P2 Allocation':[],
    'P3 Allocation':[],
    'Active?':[],
    'Weekly Performance':[],
    'Current Investment Dollar Value':[],
    'All-Time Performance':[],
}, columns=['Name','Email','SMS Email','IIDate','IIAmount','P1 Allocation','P2 Allocation','P3 Allocation','Active?','Weekly Performance','Current Investment Dollar Value','All-Time Performance'])

# Open investor workbook
invs = xlrd.open_workbook('Investors.xlsx')

print('[Collecting Investor Data...]')

'''COLLECT INVESTOR DATA'''
# This loop puts all the relevant values from the investor excel workbook into a Pandas DataFrame.
s=0 #sheet number
while True:
    try:
        inv = invs.sheet_by_index(s) # Opens an individual investor workbook
        
        namE = inv.cell_value(0,1)
        email = inv.cell_value(2,1)
        smsEmail = inv.cell_value(3,1)
        iiDate = inv.cell_value(4,1)
        iiAmount = inv.cell_value(5,1)
        p1Allocation = inv.cell_value(8,1)
        p2Allocation = inv.cell_value(9,1)
        p3Allocation = inv.cell_value(10,1)
        active = inv.cell_value(6,1)
        
        # Add investor data to the DataTable
        dfINV.loc[len(dfINV.index)] = [namE,email,smsEmail,iiDate,iiAmount,p1Allocation,p2Allocation,p3Allocation,active,'PlaceHolder','PlaceHolder','PlaceHolder']
        
        s+=1 # Moves onto the next investor sheet

    except IndexError: # Stops the collection of data when the loop runs out of investor spreadsheets to look at
        break

        

# Define dictionaries for the portfolios, their stocks and their allocations.
centraL = xlrd.open_workbook('Central Sheets.xlsx')
pfs = centraL.sheet_by_index(1) # Reference the portfolio excel sheet

p1 = {pfs.cell_value(4,0):pfs.cell_value(4,1),
      pfs.cell_value(5,0):pfs.cell_value(5,1),
      pfs.cell_value(6,0):pfs.cell_value(6,1),
      pfs.cell_value(7,0):pfs.cell_value(7,1),
      pfs.cell_value(8,0):pfs.cell_value(8,1),
      pfs.cell_value(9,0):pfs.cell_value(9,1)}

p2 = {pfs.cell_value(4,2):pfs.cell_value(4,3),
      pfs.cell_value(5,2):pfs.cell_value(5,3),
      pfs.cell_value(6,2):pfs.cell_value(6,3),
      pfs.cell_value(7,2):pfs.cell_value(7,3)}

p3 = {pfs.cell_value(4,4):pfs.cell_value(4,5),
      pfs.cell_value(5,4):pfs.cell_value(5,5),
      pfs.cell_value(6,4):pfs.cell_value(6,5),
      pfs.cell_value(7,4):pfs.cell_value(7,5)}


# DataFrames for the individual portfolios
dfP1 = pd.DataFrame({
    'Ticker':[],
    'Allocation':[],
    'Last Close Price':[],
    'Weekly Change':[],
}, columns=['Ticker','Allocation','Last Close Price','Weekly Change'])

dfP2 = pd.DataFrame({
    'Ticker':[],
    'Allocation':[],
    'Last Close Price':[],
    'Weekly Change':[],
}, columns=['Ticker','Allocation','Last Close Price','Weekly Change'])

dfP3 = pd.DataFrame({
    'Ticker':[],
    'Allocation':[],
    'Last Close Price':[],
    'Weekly Change':[],
}, columns=['Ticker','Allocation','Last Close Price','Weekly Change'])


print('[Calulating All-Time Returns...]')


'''CALCULATE GAW ALL TIME RETURN AND INVESTOR ALL TIME RETURN'''
# Pulling up the GAW central sheet and storing the inception date
GAWSheet = centraL.sheet_by_index(2)
GAWDate = GAWSheet.cell_value(0,1)

# Declaring all time value change variables for GAW
changeSinceGAWP1 = 0
changeSinceGAWP2 = 0
changeSinceGAWP3 = 0


for index in p1: # cycling through each ticker to calculate its % change in price since GAW inception
        STONKData = yf.download(index, start=GAWDate) # storing yahoo finance data since GAW inception
        time.sleep(1)
        begGAWPrice = STONKData.iloc[0,3] # price at start of time period
        endGAWPrice = STONKData.iloc[-1,3] # price at end of time period
        changeGAWPrice = (endGAWPrice-begGAWPrice)/begGAWPrice # % change in the price
        wtdChangeGAWPrice = changeGAWPrice*float(p1[index]) # Weighted by its allocation in the portfolio
        changeSinceGAWP1 += wtdChangeGAWPrice # Adds the weighted change to the total portfolio change
        # Repeat for other portfolios
for index in p2:
        STONKData = yf.download(index, start=GAWDate) 
        time.sleep(1)
        begGAWPrice = STONKData.iloc[0,3]
        endGAWPrice = STONKData.iloc[-1,3]
        changeGAWPrice = (endGAWPrice-begGAWPrice)/begGAWPrice
        wtdChangeGAWPrice = changeGAWPrice*float(p2[index])
        changeSinceGAWP2 += wtdChangeGAWPrice
for index in p3:
        STONKData = yf.download(index, start=GAWDate)
        time.sleep(1)
        begGAWPrice = STONKData.iloc[0,3]
        endGAWPrice = STONKData.iloc[-1,3]
        changeGAWPrice = (endGAWPrice-begGAWPrice)/begGAWPrice
        wtdChangeGAWPrice = changeGAWPrice*float(p3[index])
        changeSinceGAWP3 += wtdChangeGAWPrice

# calculation for all time GAW return is later on under '''FINAL CALCULATIONS'''

# Calculate and store investor all time return for this portfolio (this code is getting bad)
for g in range(len(dfINV.index)):
    
    if dfINV.iloc[g,8] == 'y':
        
        iP1Chng = 0
        iP2Chng = 0
        iP3Chng = 0

        wtIP1Chng = 0
        wtIP2Chng = 0
        wtIP3Chng = 0
        
        iInvDate = dfINV.iloc[g,3] #Initial investment date
        
        for index in p1: 
            STONKData = yf.download(index, start=iInvDate) #data from initial inv date
            time.sleep(1)
            begP = STONKData.iloc[0,3] 
            endP = STONKData.iloc[-1,3] 
            chngP = (endP-begP)/begP 
            wtChngP = chngP*float(p1[index])#weighted change
            iP1Chng += wtChngP #adds to cumulative p1 change from that investors initial investment date
        wtIP1Chng = iP1Chng*(dfINV.iloc[g,5]) #weighs the change for an overall investor change in value
            # Repeat for other portfolios      
        for index in p2: 
            STONKData = yf.download(index, start=iInvDate) #data from initial inv date
            time.sleep(1)
            begP = STONKData.iloc[0,3] 
            endP = STONKData.iloc[-1,3] 
            chngP = (endP-begP)/begP 
            wtChngP = chngP*float(p2[index])
            iP2Chng += wtChngP
        wtIP2Chng = iP2Chng*(dfINV.iloc[g,6])

        for index in p3: 
            STONKData = yf.download(index, start=iInvDate) #data from initial inv date
            time.sleep(1)
            begP = STONKData.iloc[0,3] 
            endP = STONKData.iloc[-1,3] 
            chngP = (endP-begP)/begP 
            wtChngP = chngP*float(p3[index])
            iP3Chng += wtChngP
        wtIP3Chng = iP3Chng*(dfINV.iloc[g,7])                                              

        totalInvChange = (wtIP1Chng+wtIP2Chng+wtIP3Chng)
        dfINV.iloc[g,-1] = totalInvChange
    else:
        continue

print('[Calculating Weekly Returns...]')

'''CALCULATE GAW WEEKLY RETURN AND INVESTOR WEEKLY RETURN'''
# Declaring portfolio weekly change variables
wkChangeP1 = 0
wkChangeP2 = 0
wkChangeP3 = 0


for index in p1:
    v = p1[index] # storing allocation
    stonk = yf.Ticker(index)
    time.sleep(1)
    stonkHist = stonk.history(period='5d') # five day (weekly) historical data
    begWkPrice = stonkHist.iloc[0,3] # close price on beg of week
    endWkPrice = stonkHist.iloc[-1,3] # close price on end of week
    changeWkPrice = (endWkPrice-begWkPrice)/begWkPrice
    wtChangeWkPrice = changeWkPrice*float(v)
    wkChangeP1 += wtChangeWkPrice
    # Add portfolio data to the DataTable
    dfP1.loc[len(dfP1.index)] = [index,v,endWkPrice,changeWkPrice]
    
for index in p2:
    v = p2[index]
    stonk = yf.Ticker(index)
    time.sleep(1)
    stonkHist = stonk.history(period='5d')
    begWkPrice = stonkHist.iloc[0,3] 
    endWkPrice = stonkHist.iloc[-1,3]
    changeWkPrice = (endWkPrice-begWkPrice)/begWkPrice
    wtChangeWkPrice = changeWkPrice*float(v)
    wkChangeP2 += wtChangeWkPrice
    dfP2.loc[len(dfP2.index)] = [index,v,endWkPrice,changeWkPrice]
    
for index in p3:
    v = p3[index] 
    stonk = yf.Ticker(index)
    time.sleep(1)
    stonkHist = stonk.history(period='5d') 
    begWkPrice = stonkHist.iloc[0,3] 
    endWkPrice = stonkHist.iloc[-1,3] 
    changeWkPrice = (endWkPrice-begWkPrice)/begWkPrice
    wtChangeWkPrice = changeWkPrice*float(v)
    wkChangeP3 += wtChangeWkPrice
    dfP3.loc[len(dfP3.index)] = [index,v,endWkPrice,changeWkPrice]

# calculation for weekly gaw return is later on under '''FINAL CALCULATIONS'''

# Calculate weekly change in value for each investor
for g in range(len(dfINV.index)):
    if dfINV.iloc[g,8] == 'y':
        p1WkChng = dfP1.iloc[0,3]
        wtP1WkChng = p1WkChng*float(dfINV.iloc[g,5])
        p2WkChng = dfP2.iloc[0,3]
        wtP2WkChng = p2WkChng*float(dfINV.iloc[g,6])
        p3WkChng = dfP3.iloc[0,3]
        wtP3WkChng = p3WkChng*float(dfINV.iloc[g,7])

        wkChngINV = (wtP1WkChng+wtP2WkChng+wtP3WkChng)

        dfINV.iloc[g,-3] = wkChngINV
        
    else:
        continue

print('[Calculating Current Values...]')

'''CALCULATE CURRENT INVESTMENT DOLLAR VALUE AND MANAGED FUNDS VALUE'''
managedFundsValue = 0 
for g in range(len(dfINV.index)):
    if dfINV.iloc[g,8] == 'y':
        chng = float(dfINV.iloc[g,-1]) # % change in value
        begValue = dfINV.iloc[g,4]
        curIValue = begValue*(1+chng)
        dfINV.iloc[g,-2] = curIValue # Assigning to dataframe
        managedFundsValue += curIValue # Adding to total managed funds value
    else:
        continue
    

print('[Finalizing Calculations...]')

'''FINAL CALCULATIONS'''
# Calculates total GAW change in value since inception from all of the portfolios
totalGAWChange = (changeSinceGAWP1+changeSinceGAWP2+changeSinceGAWP3)

# Calculates weekly GAW change in value from all of the portfolios
weeklyGAWChange = (wkChangeP1+wkChangeP2+wkChangeP3)


print('[Sending Emails...]')

'''SEND MESSAGES'''
# email/server setup
password = 'placeholder'
port = 465 # For SSL
senderEmail = 'placeholder'
message = MIMEMultipart('alternative')
message['Subject'] = 'Weekly GAW Investor Report'

for g in range(len(dfINV.index)):
    if dfINV.iloc[g,8] == 'y':
        
        name = dfINV.iloc[g,0]
        whosEmail = dfINV.iloc[g,1] #investor email address

        val2 = str((dfINV.iloc[g,-3])*100) # investor weekly return
        val2 = val2[0:6]#this shortens the string bc its so long, i know its a bad fix.
        val3 = str((dfINV.iloc[g,-1])*100) # investor all time return
        val3 = val3[0:6]
        val4 = str(dfINV.iloc[g,4]) # investor initial value
        val5 = str(dfINV.iloc[g,-2]) # investor current value
        val5 = val5[0:6]
        val6 = str((weeklyGAWChange)*100) # weekly GAW return
        val6 = val6[0:6]
        val7 = str((totalGAWChange)*100) # total GAW return
        val7 = val7[0:6]
        val8 = str(managedFundsValue) # total managed value
        val8 = val8[0:6]

        message['From'] = senderEmail
        message['To'] = whosEmail

        # Plain text message
        text = '''\
        Good afternoon %s,

        Here is your weekly update on your investment with GAW.

        Weekly Return: %s %%
        All-Time Return: %s %%
        Initial Value: $ %s
        Current Value: $ %s

        [GAW STATS]
        Weekly Return: %s %%
        All-Time Return: %s %%
        Managed Funds Value: $ %s

        Thank you for investing with Guh Associated Wealth, and remember that these numbers are only estimates.
        Good night king.
        ''' % (name, val2, val3, val4, val5, val6, val7, val8)

        # HTML version
        html = '''\
        <html>
            <body>
                <h1>
                Good afternoon %s,
                </h1>
                <h1>
                Here is your weekly update on your investment with GAW.
                </h1>
                
                <p> </p>
                
                <p>
                Weekly Return: %s %%
                </p>
                <p>
                All-Time Return: %s %%
                </p>
                <p>
                Initial Value: $ %s
                </p>
                <p>
                Current Value: $ %s
                </p>
                
                <p> </p>

                <p>
                [GAW STATS]
                </p>
                <p>
                Weekly Return: %s %%
                </p>
                <p>
                All-Time Return: %s %%
                </p>
                <p>
                Managed Funds Value: $ %s
                </p>

                <p> </p>

                <p>
                Thank you for investing with Guh Associated Wealth, and remember that these numbers are only estimates.
                Good night king.
                </p>
            </body>
        </html>
                ''' % (name, val2, val3, val4, val5, val6, val7, val8)

        part1 = MIMEText(text,'plain')
        part2 = MIMEText(html,'html')

        message.attach(part1)
        message.attach(part2)

        context = ssl.create_default_context() # Create a secure SSL context
        with smtplib.SMTP_SSL('smtp.gmail.com', port, context=context) as server:
            server.login(senderEmail, password)
            server.sendmail(senderEmail, whosEmail, message.as_string()) # Send email

            pPrint = '''[Email Sent To %s...]''' % (name)
            print(pPrint)
        
    else:
        continue

print('[All Done! Good Night King...]')

time.sleep(5)
sys.exit()
