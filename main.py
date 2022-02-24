# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
# open account and parse email
import imaplib
import email
import os
import time
import codecs
import pandas as pd
from bs4 import BeautifulSoup
import re
import datetime
import calendar
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
def findDay(date):
    born = datetime.datetime.strptime(date, '%d-%m-%Y').weekday()
    return (calendar.day_name[born])

username ="tradexdata25@gmail.com"
app_password= "dvhypvwysosbprps"
#tznurcdwlnkkqmou
gmail_host= 'imap.gmail.com'
mail = imaplib.IMAP4_SSL(gmail_host)
mail.login(username, app_password)
# tradedata
mail.select("tradedata")
T=time.time()
abc, selected_mails = mail.search(None, '(UNSEEN)')
sheetlist=[]

a={}
my_file = open("users.txt", "r")
users = my_file.read().split("/n")
my_file.close()
j=1
#now from here we can continue
messagedict={}


for num in selected_mails[0].split():
    abc, data = mail.fetch(num , '(RFC822)')
    abc, bytes_data = data[0]
    full_data=[]
    mainlist=[]
    clientcode=""
    email_message = email.message_from_bytes(bytes_data)

    #access data
    if ("CONTRACT NOTE F&O" in email_message["subject"]):
        
        subject_list=email_message["subject"].split()
        if " " in subject_list:
            subject_list.remove(' ')
        if None in subject_list:
            subject_list.remove(None)
        ind=subject_list.index('F&O')
        clientcode=subject_list[ind+1]
    else:
        continue
    for part in email_message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()
        if bool(fileName) and '.html' in fileName:
            fl=" "
            fl=part.get_payload(decode=True)
#             if not os.path.isfile(filePath) :
#                 fp = open(filePath, 'wb')
#                 part.get_payload(decode=True)
#                 fp.close()
#             HTMLFileToBeOpened = open(filePath, "r")
#             contents = HTMLFileToBeOpened.read() 
            beautifulSoupText = BeautifulSoup(fl, 'lxml')
            #HTMLFileToBeOpened.close()
            table = beautifulSoupText.findAll("table")[8]
            for row in table.find_all('tr')[1:]:
                row_data=[]
                for col in row.find_all('td'):
                    row_data.append(col.text)
                full_data.append(row_data)
            appenddata=[]
            niftycebuytime=""
            niftypebuytime=""
            bankniftycebuytime=""
            bankniftypebuytime=""
            bankniftyce=0
            bankniftyceS=0
            bankniftyceB=0
            bankniftyqty=0
            niftyqty=0
            bankniftytradetime=""
            niftytradetime=""
            bankniftype=0
            bankniftypeS=0
            bankniftypeB=0
            bankniftystrike=""
            bankniftystrikeback=""
            niftystrike=""
            niftystrikeback=""
            niftyce=0
            niftype=0
            niftyceS=0
            niftypeS=0
            niftyceB=0
            niftypeB=0
            niftypnl=0
            bankniftypnl=0
            date = beautifulSoupText.findAll("table")[5].findAll('td')[5].text.strip()
            tables = beautifulSoupText.findAll("table")[10]
            payout= tables.find_all('tr')[1].find_all('td',{"align":"right"})[0].text
            netamount= tables.find_all('tr')[9].find_all('td',{"align":"right"})[0].text
            expenses= abs(float(netamount)-float(payout))
            
            for count,i in enumerate(full_data):
                if (i[4]!="Scrip Total "):
                    if ("BANKNIFTY"in i[4] and "CE" in i[4]):
                        if (i[5]=="S"):
                            bankniftyceS=bankniftyceS+float(i[10].strip())
                        elif (i[5]=="B"):
                            bankniftyceB=bankniftyceB+abs(float(i[10].strip()))
                        if(bankniftyce==0):
                            bankniftytradetime=i[3]
                            bankniftystrike=re.search('.+-(.+)-.+', i[4])
                            bankniftystrike=float(bankniftystrike.group(1))

                            bankniftyce=1
                        if(bankniftycebuytime=="" and i[5]=="B"):
                            bankniftycebuytime=i[3]
                            
                    elif ("BANKNIFTY"in i[4] and "PE" in i[4]):
                        if (i[5]=="S"):
                            bankniftypeS=bankniftypeS+float(i[10].strip())
                        elif (i[5]=="B"):
                            bankniftypeB=bankniftypeB+abs(float(i[10].strip()))
                        if(bankniftype==0):
                            bankniftytradetime=i[3]
                            bankniftype=1
                            bankniftystrike=re.search('.+-(.+)-.+', i[4])
                            bankniftystrike=float(bankniftystrike.group(1))
                        if(bankniftypebuytime=="" and i[5]=="B"):
                            bankniftypebuytime=i[3]
                    elif ("NIFTY"in i[4] and "CE" in i[4]):
                        if (i[5]=="S"):
                            niftyceS=niftyceS+float(i[10].strip())
                        elif (i[5]=="B"):
                            niftyceB=niftyceB+abs(float(i[10].strip()))
                        if(niftyce==0):
                            niftytradetime=i[3]
                            niftystrike=re.search('.+-(.+)-.+', i[4])
                            niftystrike=float(niftystrike.group(1))
                            niftyce=1
                        if(niftycebuytime=="" and i[5]=="B"):
                            niftycebuytime=i[3]
                    elif ("NIFTY"in i[4] and "PE" in i[4]):
                        if (i[5]=="S"):
                            niftypeS=niftypeS+float(i[10].strip())
                        elif (i[5]=="B"):
                            niftypeB=niftypeB+abs(float(i[10].strip()))
                        if(niftype==0):
                            niftytradetime=i[3]
                            niftystrike=re.search('.+-(.+)-.+', i[4])
                            niftystrike=float(niftystrike.group(1))
                            niftype=1
                        if(niftypebuytime=="" and i[5]=="B"):
                            niftypebuytime=i[3]
                elif (i[4]=="Scrip Total "):
                    qty=float(i[5].strip())
                    if((bankniftyce!=0 or bankniftype!=0)and( count==len(full_data)-1 or str(bankniftystrike) not in re.search('.+-(.+)-.+',full_data[count+1][4]).group(1))):
                        bankniftypnl=bankniftypnl+float(i[10].strip())
                        alist=[] 
                        alist.append(findDay(date)[:3])
                        alist.append(date)
                        alist.append("BankNifty")
                        alist.append(qty)
                        alist.append(bankniftypnl)
                        alist.append(bankniftytradetime)
                        alist.append(bankniftystrike)
                        if bankniftyceS!=0:
                            alist.append(round(bankniftyceS/qty,2))
                        else:
                            alist.append(0)
                        if bankniftypeS!=0:
                            alist.append(round(bankniftypeS/qty,2))
                        else:
                            alist.append(0)
                        if bankniftyceB!=0:
                            alist.append(round(bankniftyceB/qty,2))
                        else:
                            alist.append(0)
                        if bankniftypeB!=0:
                            alist.append(round(bankniftypeB/qty,2))
                        else:
                            alist.append(0)
                        alist.append(" ")
                        var=0
                        if (bankniftyceB>bankniftyceS and bankniftycebuytime<'15:19:00'):
                            if var==0:
                                var ='CE'
                            else:
                                var=var+'+PE'
                        if (bankniftypeB>bankniftypeS and bankniftypebuytime<'15:19:00'):
                            if var==0:
                                var ='PE'
                            else:
                                var=var+'+PE'
                        alist.append(var)       
                        #sheet_instance.append_row(alist)
                        appenddata.append(alist)
                        bankniftyce=0
                        bankniftyceS=0
                        bankniftyceB=0
                        bankniftyqty=0
                        bankniftytradetime=""
                        bankniftype=0
                        bankniftypeS=0
                        bankniftypeB=0
                        bankniftypnl=0
                    elif (bankniftyce!=0 or bankniftype!=0):
                        bankniftypnl=float(i[10].strip())
                        
                    elif ((niftyce!=0 or niftype!=0)and(count==len(full_data)-1 or str(niftystrike) not in re.search('.+-(.+)-.+', full_data[count+1][4]).group(1))):
                        niftypnl=niftypnl+float(i[10].strip())
                        blist=[] 
                        blist.append(findDay(date)[:3])
                        blist.append(date)
                        blist.append("Nifty")
                        blist.append(qty)
                        blist.append(niftypnl)
                        blist.append(niftytradetime)
                        blist.append( niftystrike)
                        if niftyceS!=0:
                            blist.append(round(niftyceS/qty,2))
                        else:
                            blist.append(0)
                        if niftypeS!=0:
                            blist.append(round(niftypeS/qty,2))
                        else:
                            blist.append(0)
                        if niftyceB!=0:
                            blist.append(round(niftyceB/qty,2))
                        else:
                            blist.append(0)
                        if niftypeB!=0:
                            blist.append(round(niftypeB/qty,2))
                        else:
                            blist.append(0)
                        blist.append(" ")
                        var=0
                        if (niftyceB>niftyceS and niftycebuytime<'15:19:00'):
                            if var==0:
                                var ='CE'
                            else:
                                var=var+'+PE'
                        if (niftypeB>niftypeS and niftypebuytime<'15:19:00'):
                            if var==0:
                                var ='PE'
                            else :
                                var=var+'+PE'
                        blist.append(var)
                        #sheet_instance.append_row(blist)
                        appenddata.append(blist)
                        niftyce=0
                        niftype=0
                        niftyce=0
                        niftype=0
                        niftyceS=0
                        niftyqty=0
                        niftypeS=0
                        niftyceB=0
                        niftypeB=0
                        niftypnl=0
                    elif (niftyce!=0 or niftype!=0):
                        niftypnl=float(i[10].strip())
            appenddata[-1][-2]=round(expenses,2)
            if clientcode in messagedict.keys():
                messagedict[clientcode]=messagedict[clientcode]+appenddata
            else :
                messagedict[clientcode]=appenddata
            subject = str(email_message).split("Subject: ", 1)[1].split("\nTo:", 1)[0]
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
# add credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name('tradex-342202-54c8c7955276.json', scope)
# authorize the clientsheet 
client = gspread.authorize(creds)

for i in messagedict:
    messagedict[i]=sorted(messagedict[i], key = lambda x: datetime.datetime.strptime(x[1], '%d-%m-%Y'))
    if i in users:

        sheet = client.open(i)
        sheet_instance = sheet.worksheet("data")
        sheet_instance.append_rows(messagedict[i]) 

    else:
        sh1 =client.create(i)
        sh1.share('tradexdata25@gmail.com', perm_type='user', role='owner')
        worksheet = sh1.add_worksheet(title="data",rows="100", cols="20")
        sheet_instance = sh1.worksheet("data")
        sheet_instance.append_row(["Trade Type",	"Date",	"Instrument",	"Qty",	"PnL",	"Entry Time",	"Strike",	"CE Sell Price",	"PE Sell Price",	"CE Buy Price",	"PE Buy Price",	"Expense",	"SL HIT"])
        worksheet.format('A1:M1', {'textFormat': {'bold': True}})
        sheet_instance.append_rows(messagedict[i])

        my_file = open("users.txt", "a")
        my_file.write(i)
        my_file.close()
        users.append(i)


        

            
