# dm-notices-
import tempfile
from docx import Document
import docx
import sys
import os
import io  # interfaces to stream handling
from datetime import date
import win32com.client as win32
import time

tempDocLoc = ""
today = date.today()
print today
x = ' '

print '\n'
print "For assistance with using this program please e-mail sahdique.mohamed@aucklandcouncil.govt.nz"
print '\n'
print "****Interview for Manager's Certificate Generator***"
print '\n'
print "Please follow the instructions below"
print '\n'

def email(file_to_attach):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = outlookadd
    mail.Subject = "Manager's Certificate Application - " + str(TID)
    mail.Body = 'Kia Ora \n \nYour interview is scheduled as follows: \n \nDate: ' + interviewdate + "\nTime: " + interviewtime + '\nLocation: Henderson Service Centre - Level 2/6 Henderson Valley Road, Henderson, Auckland \n\nPlease bring along a print out of the attached PDF. \n\nEnsure that you have read through the document carefully, as it contains information that will help you in your preparation for the interview. \n\nIf you need further assitance please contact me.'
    mail.Attachments.Add(Source=file_to_attach)
    mail.Display(True)

def start():
    f = open('BASE.docx', 'rb')  #rb = reading in binary
    global templateDoc
    templateDoc = Document(f) #using base.docx
    f.close() #good practice to close
    global Name
    Name = raw_input("Insert applicant's name here: ")
    global TID
    TID = raw_input('Insert Transaction No. here: ')
    global interviewdate
    interviewdate = raw_input("Insert applicant's interview date here: ")
    global interviewtime
    interviewtime = raw_input("Insert applicant's interview time here: ")
    global outlookadd
    outlookadd = raw_input("Insert applicant's e-mail here: ")

def formatdate():
    today = str(date.today())
    year = today[:4]
    month = today[5:7]
    day = today[8:]
    months = ['index', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
              'November', 'December']
    dmonth = ''
    if month == '00':
        dmonth += months[0]
    elif month == '01':
        dmonth += months[1]
    elif month == '02':
        dmonth += months[2]
    elif month == '03':
        dmonth += months[3]
    elif month == '04':
        dmonth += months[4]
    elif month == '05':
        dmonth += months[5]
    elif month == '06':
        dmonth += months[6]
    elif month == '07':
        dmonth += months[7]
    elif month == '08':
        dmonth += months[8]
    elif month == '09':
        dmonth += months[9]
    elif month == '10':
        dmonth += months[10]
    elif month == '11':
        dmonth += months[11]
    elif month == '12':
        dmonth += months[12]
    today = day + x + dmonth + x + year
    return today


def today_date():
    for paragraph in templateDoc.paragraphs:
        if 'CDate' in paragraph.text:
            paragraph.text = formatdate()


def applicant_name():
    for paragraph in templateDoc.paragraphs:
        if 'NAME' in paragraph.text:
            paragraph.text = 'Kia Ora ' + Name


def tranID():
    for paragraph in templateDoc.paragraphs:
        if 'TID' in paragraph.text:
            paragraph.text = 96 * x + 'Transaction No. ' + TID


def targetdate():
    for paragraph in templateDoc.paragraphs:
        if 'IDATE' in paragraph.text:
            paragraph.text = 3 * x + 'Date:' + 30 * x + interviewdate


def targettime():
    for paragraph in templateDoc.paragraphs:
        if 'ITIME' in paragraph.text:
            paragraph.text = 3 * x + 'Time:' + 30 * x + interviewtime


def initialize():
    current = 0
    target = 100
    while current < target:
        start()
        today_date()
        applicant_name()
        tranID()
        targetdate()
        targettime()
        tempDocLoc = tempfile.gettempdir() + "\\Manager's Interview - " + TID + " " + Name + '.docx' #templatedocument location
        templateDoc.save(tempDocLoc)
        print 'Generating interview document'
        print '\n'
        time.sleep(0.5)
        print '!!!! ' + Name + 's generation successful!!!!'
        print '\n'
        print '\n'
        tempDocLoc.replace('\\', r'\\') #needs double back slash - kind of like a security feature
        print 'Check your Outlook!'
        email(tempDocLoc)
        current += 1


initialize()
