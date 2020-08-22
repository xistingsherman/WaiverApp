import tkinter as tk
from tkinter import font  as tkfont
import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import re
from tkcalendar import Calendar, DateEntry
from bs4 import BeautifulSoup
import sys
import traceback
import datetime
import copy
from datetime import date
from calendar import timegm

#import email.header
#import datetime

class WaiverApp(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        #https://stackoverflow.com/questions/7546050/switch-between-two-frames-in-tkinter
        self.title_font = tkfont.Font(family='Helvetica', size=18, weight="bold", slant="italic")
        self.geometry('300x300')
        # the container is where we'll stack a bunch of frames
        # on top of each other, then the one we want visible
        # will be raised above the others
        container = tk.Frame(self)
        self.resizable(width=False, height=False)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        #Must turn on less secure app in security https://myaccount.google.com/security
        #code taken from https://www.thepythoncode.com/article/reading-emails-in-python

        self.password = ""
        self.numOfMessages = tk.StringVar()
        self.numOfMessages.set('0')
        self.intOfMessages = 0
        self.messages = ""
        self.list = []
        self.imap = ""
        self.username = "waivers@rosebowlaquatics.org"
        #container.imap

        self.frames = {}
        for F in (StartPage, PageOne, PageTwo, PageThree):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame

            # put all of the pages in the same location;
            # the one on the top of the stacking order
            # will be the one that is visible.
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("StartPage")

    def show_frame(self, page_name):
        '''Show a frame for the given page name'''
        frame = self.frames[page_name]
        frame.tkraise()

#http://tech.franzone.blog/2012/11/24/listing-imap-mailboxes-with-python/
    def connectToServer(self):
        # account credentials
        try:
            # Create the IMAP Client
            #imap = imaplib.IMAP4(host, port)
            self.imap = imaplib.IMAP4_SSL("imap.gmail.com")
            # Login to the IMAP server
            resp, data = self.imap.login(self.username, self.password)
            status, self.messages = self.imap.select("INBOX") #change this to optional


            if resp == 'OK':
                # List all mailboxes
                resp, data = self.imap.list('""', '*')
                if resp == 'OK':
                    for mbox in data:
                        flags, separator, name = self.parse_mailbox(bytes.decode(mbox))
                        fmt = '{0}    : [Flags = {1}; Separator = {2}'
                        fmt = '{0}'
                        if not "/ [Gmail]" in fmt.format(name):
                            temp = tk.StringVar()
                            temp.set(fmt.format(name))
                            self.list.append(temp.get())
                            #self.list.append(fmt.format(name))
                            #list.extend(fmt.format(name))
        except:
            print("Temporarily allow less secure apps to connect with email.")
            return

        self.imap.close()
        self.imap.logout()
        #for x in self.list:
            #print(x.get())
        self.numOfMessages.set(self.messages[0]) #left off here 7/1/2020
        self.intOfMessages = int(self.messages[0])
        #print(self.intOfMessages)
        #print(messages[0])
        self.show_frame("PageOne")

    def parse_mailbox(self,data):
        flags, b, c = data.partition(' ')
        separator, b, name = c.partition(' ')
        return (flags, separator.replace('"', ''), name.replace('"', ''))

    #get Waivers by number
    def get_waivers_by_number(self,N,mailbox):
        try:
            text_file = open("Output.csv", "w")

            self.imap = imaplib.IMAP4_SSL("imap.gmail.com")
            resp, data = self.imap.login(self.username, self.password)
            try:
                status, self.messages = self.imap.select(mailbox)
            except:
                status, self.messages = self.imap.select("INBOX")
                print("Error: Invalid inbox chosen. Desired inbox may be empty. Selecting 'Inbox'.")

            for i in range(self.intOfMessages, self.intOfMessages-N, -1):
                res, msg = self.imap.fetch(str(i), "(RFC822)")
                for response in msg:
                    if isinstance(response, tuple):
                        # parse a bytes email into a message object
                        msg = email.message_from_bytes(response[1])
                        # decode the email subject
                        subject = decode_header(msg["Subject"])[0][0]
                        if isinstance(subject, bytes):
                            # if it's a bytes, decode to str
                            subject = subject.decode()

                        emailDate = msg.get("Date")
                        emailDate = timegm(email.utils.parsedate_tz(emailDate))
                        emailDate = datetime.datetime.fromtimestamp(emailDate)
                        # if the email message is multipart
                        if msg.is_multipart():
                            # iterate over email parts
                            for part in msg.walk():
                                # extract content type of email
                                content_type = part.get_content_type()
                                content_disposition = str(part.get("Content-Disposition"))

                                try:
                                    break
                                except:
                                    pass
                                if content_type == "text/plain" and "attachment" not in content_disposition:
                                    break
                        else:
                            content_type = msg.get_content_type()
                            try:
                                body = msg.get_payload(decode=True).decode('UTF-8') #does not take accented characters
                            except:
                                body = msg.get_payload(decode=True).decode('ISO-8859-1')

                            if content_type == "text/plain":
                                print(body)
                        if content_type == "text/html" and subject == "Form 'RBAC Waiver of Liability' Submission Received":
                            self.filter_text(body,text_file,emailDate)

                        print("="*100)
                    else:
                        break
            text_file.close()
            self.imap.close()
            self.imap.logout()
            self.show_frame("PageTwo")
        except OSError as e:
            if e.errno == 13:
                print('Output.csv file is open. Please close file and rerun program.')
        except UnboundLocalError as error:
            # Output expected UnboundLocalErrors.
            print('Please restart program.')

    def get_waivers_by_date(self,startDate,endDate,mailbox):
        try:
            text_file = open("Output.csv", "w")

            self.imap = imaplib.IMAP4_SSL("imap.gmail.com")
            resp, data = self.imap.login(self.username, self.password)
            try:
                status, self.messages = self.imap.select(mailbox)
            except:
                status, self.messages = self.imap.select("INBOX")
                print("Error: Invalid inbox chosen. Desired inbox may be empty. Selecting inbox")

            today = date.today()
            start = startDate.strftime("%d-%b-%Y")
            end = endDate.strftime("%d-%b-%Y")

            if startDate < today and endDate >= today:
                response, message = self.imap.search(None, '(SINCE "' + start + '")')
            else:
                response, message = self.imap.search(None, '(SINCE "' + start + '" BEFORE "' + end + '")')

                #https://gist.github.com/robulouski/7441883

            if response != 'OK':
                print ("No messages found!")
                return

            for each in message[0].split():
                rev, data = self.imap.fetch(each, "(RFC822)")
                if rev != 'OK':
                    print("ERROR getting message " + each)
                    return
                msg = email.message_from_string(data[0][1])
                decode = email.header.decode_header(msg['Subject'])[0]
                subject = unicode(decode[0])

                print(message)
                print(decode)
                print(subject)

                emailDate = email.header.decode_header(msg['Date'])[0]
                print(emailDate)
                emailDate = timegm(email.utils.parsedate_tz(emailDate))
                emailDate = datetime.datetime.fromtimestamp(emailDate)
                currentDate = emailDate.date()
                        #print(date)
                        # if the email message is multipart
                if msg.is_multipart():
                            # iterate over email parts
                    for part in msg.walk():
                                # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))

                        try:
                            break
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            break
                else:
                    content_type = msg.get_content_type()
                    try:
                        body = msg.get_payload(decode=True).decode('UTF-8') #does not take accented characters
                    except:
                        body = msg.get_payload(decode=True).decode('ISO-8859-1')

                    if content_type == "text/plain":
                        print(body)
                if content_type == "text/html" and subject == "Form 'RBAC Waiver of Liability' Submission Received" and currentDate <= endDate and currentDate >= startDate:
                    self.filter_text(body,text_file,emailDate)

                    print("="*100)
                else:
                    break
        finally:
            text_file.close()
            self.imap.close()
            self.imap.logout()
            self.show_frame("PageTwo")

    def filter_text(self,body,text_file,emailDate):
        w1 = "You have received a completed RBAC Waiver of Liability submission. Please see below details. You can also access this through the online portal."
        w2 = "Form Submission"
        w3 = "Rose Bowl Aquatics Center"
        w4 = "RBAC Waiver of Liability"
        w5 = "Type :"
        w6 = "Form"
        w7 = "User Type :"
        w8 = "Kiosk"
        w9 = "Name :"
        w10 = "rbac_website"
        w11 = "Link :"
        w12 = "Form"
        w13 = "Instructions :"
        w14 = "Questions"
        w15 = "Response"
        w16 = "In consideration of the ROSE BOWL AQUATICS CENTER (“RBAC”) granting me, or any minor on whose behalf I sign this agreement (“collectively, the Participant”), permission to use facilities leased to RBAC and/or participate in any RBAC-sponsored activities on or off-site, the undersigned voluntarily agrees to the following contractual terms and conditions:"
        w17 = "1. GOOD PHYSICAL CONDITION. The Participant has no physical or medical condition which would endanger the Participant or others, or that would interfere with the Participant’s ability to participate in the event/program."
        w18 = "2. ALL RISKS ASSUMED. I am fully aware that serious injuries and possibly even death are sometimes associated with such event/programs, and sporting and recreational activities. I fully realize the dangers and hazards associated with participating in this event/program and fully assume the risks associated with such participation, including, by way of example and not limitation, the following: dangers of falling, tripping, drowning, hitting (coming into physical contact) or being hit by other participants, spectators and fixed or moving objects; dangers arising from facility defects or surface hazards, equipment failure or lack thereof; inadequate safety equipment; and weather conditions; and exposure to communicable diseases, including, but not limited to COVID-19."

        w20 = "I accept responsibility to be familiar with the premises, the equipment, the improvements, the weather, and the rules and practices regulating the event/program. Knowing the risks and dangers, I nevertheless agree to assume, for myself and Participant, all event/program risks and dangers (known and unknown, foreseen and unforeseen, and whether mentioned above or not)."
        w21 = "3. RELEASE, WAIVER, INDEMNITY, AND COVENANT NOT TO SUE. I agree for myself, or any minor on whose behalf I sign this agreement, and for our executors, administrators, heirs, next of kin, successors and assigns (collectively hereafter called successors) to waive, release, discharge, agree not to sue, and agree to indemnify, hold harmless and defend, to the extent permitted by law, RBAC, its respective directors, officers. employees, volunteers and agents from any and all liability, loss, suits, claims, damages, costs, judgments, and expenses, including attorney’s fees and costs of litigation, which directly or indirectly result from or arise in any way out of, or are alleged to result from or arise out of, participation or association with the event/program, including, but not limited to, personal injury, (including death at any time) and property damage or other damage sustained by me or the minor participant, on whose behalf I am signing this agreement, or any person or persons whatsoever, from any cause whatsoever, whether caused by negligence or not. This release is intended to discharge in advance, RBAC, its respective directors, officers, employees, volunteers and agents from and against all liability arising out of or, in any way, connected with me or my child’s participation in said program, even if that liability may arise out of negligence or carelessness on the part of RBAC, its respective directors, officers, employees, volunteers and agents."
        w22 = "4. MEDICAL AUTHORIZATION/CONSENT FOR MEDICAL TREATMENT. I agree that this release applies to persons or entities rendering emergency medical treatment. I hereby consent that I, or any minor participant on whose behalf I am signing, may receive emergency medical treatment that may be deemed advisable in the event of injury, accident and/or illness during any program or event at RBAC or any other facility where RBAC may be conducting programs. This authorization is given in advance of any specific diagnosis, treatment, or medical care being required, and pursuant to the provisions of California Family Code Section 6900 et seq. In the event RBAC is unable to contact me or to secure my consent in the case of a medical emergency involving my child, I hereby give RBAC and its representatives permission to secure proper medical care and assistance for my child, including, but not limited to, hospitalization, treatment, medication, or x-rays. I further authorize any treating physician to use his or her discretion in providing emergency treatment. I agree to pay the costs of all such medical care."
        w23 = "5. PHOTO MEDIA RELEASE, I hereby give my consent for the use of any photographs/videos taken of me, or any minor participant on whose behalf I am signing, for such publicity as RBAC chooses and release all claims whatsoever which may arise in said regard."
        w24 = "I HAVE READ THIS ENTIRE DOCUMENT. I UNDERSTAND IT IS A RELEASE OF ALL CLAIMS AND THAT IT IS WRITTEN TO BE AS BROAD AND INCLUSIVE AS LEGALLY PERMITTED BY THE STATE OF CALIFORNIA. I UNDERSTAND AND I ASSUME ALL RISKS OF INJURY INVOLVED IN THESE ACTIVITIES AND VOLUNTARILY SIGN MY NAME FOR MYSELF, AND/OR ANY PARTICIPATING MINOR CHILD."
        w25 = "PLEASE SIGN AND DATE."
        w26 = "wcontent"
        w27 = "Name of Participant #1"

        w28 = "Name of Participant #2"
        w29 = "Name of Participant #3"
        w30 = "Name of Participant #4"
        w31 = "Name of Participant #5"
        w32 = "Parent or Guardian Name (if signing for a minor)"
        w33 = "Signature of Participant or Parent/Guardian"
        w34 = "True"
        w35 = "Date Signed"
        w19 = "© 2020 Connect2Concepts LLC.All rights reserved.Version "



        soup = BeautifulSoup(body, "html.parser" )
        body = soup.get_text('\n')
        body = body.replace("\n", "")
        body = body.replace(w1,'\n')
        body = body.replace(w2,'')
        body = body.replace(w3,'')
        body = body.replace(w4,'')
        body = body.replace(w5,'')
        body = body.replace(w6,'')
        body = body.replace(w7,'')
        body = body.replace(w8,'')
        body = body.replace(w9,'')
        body = body.replace(w10,'')
        body = body.replace(w11,'')
        body = body.replace(w12,'')
        body = body.replace(w13,'')
        body = body.replace(w14,'')
        body = body.replace(w15,'')
        body = body.replace(w16,'')
        body = body.replace(w17,'')
        body = body.replace(w18,'')
        body = body.replace(w20,'')
        body = body.replace(w21,'')
        body = body.replace(w22,'')
        body = body.replace(w23,'')
        body = body.replace(w24,'')
        body = body.replace(w25,'')
        body = body.replace(w26,'')
        body = body.replace(w27,'')
        body = body.replace(w33,'')
        body = body.replace(w34,'')
        body = body.replace(w19,',')
        body = body.replace("User",'')
        body = body.replace("content",'')

        body = body.replace(w28,',')
        body = body.replace(w29,',')
        body = body.replace(w30,',')
        body = body.replace(w31,',')
        body = body.replace(w32,',')
        body = body.replace(w35,',')

        body = body.strip()

        list = body.split(",")
        print(list)
        participant0 = copy.deepcopy(list[0])
        participant1 = copy.deepcopy(list[1])
        participant2 = copy.deepcopy(list[2])
        participant3 = copy.deepcopy(list[3])
        participant4 = copy.deepcopy(list[4])
        guardian = copy.deepcopy(list[5])
        date = copy.deepcopy(list[6])

        if (participant0 == participant1):
            participant1 = ""
        if (participant0 == participant2):
            participant2 = ""
        if (participant0 == participant3):
            participant3 = ""
        if (participant0 == participant4):
            participant4 = ""

        if (participant1 == participant2):
            participant2 = ""
        if (participant1 == participant3):
            participant3 = ""
        if (participant1 == participant4):
            participant4 = ""

        if (participant2 == participant3):
            participant3 = ""
        if (participant2 == participant4):
            participant4 = ""

        if (participant3 == participant4):
            participant4 = ""

        if (len(participant0) > 0):
            text_file.write(participant0+ ","+guardian+","+date)
            text_file.write("\n")

        if (len(participant1)> 0):
            text_file.write(participant1 + ","+guardian+","+date)
            text_file.write("\n")

        if (len(participant2) > 0):
            text_file.write(participant2 + ","+guardian+","+date)
            text_file.write("\n")

        if (len(participant3) > 0):
            text_file.write(participant3 + ","+guardian+","+date)
            text_file.write("\n")

        if (len(participant4) > 0):
            text_file.write(participant4 + ","+guardian+","+date)
            text_file.write("\n")

        #text_file.write(body)
        #text_file.write("," + emailDate.strftime("%x %X"))
        #text_file.write("\n")
        print(body)


class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text="This is the start page", font=controller.title_font)

        tk.Label(self,text='RBAC Waivers', width=15,pady=30).grid(row=0,sticky='E'+'W',pady=5, columnspan=2)
        tk.Label(self,text='Email:', width=15).grid(row=1,column=0,sticky='E',pady=5)
        #tk.Label(self,text='waivers@rosebowlaquatics.org', width=25).grid(row=1,column=1,pady=5,sticky='E')
        tk.Label(self,text='Password:', width=15).grid(row=2,column=0, sticky='W',pady=5)

        self.email_entry = tk.Entry(self,width=28)
        self.email_entry.grid(row=1, column=1,sticky='W')
        self.email_entry.insert(0, "waivers@rosebowlaquatics.org")

        self.password_entry = tk.Entry(self,width=28,show="*")
        self.password_entry.grid(row=2, column=1,sticky='W')
        self.password_entry.focus_set()
        self.controller.bind('<Return>',(lambda event: self.getPassword()))
        login_button = tk.Button(self,text="Login", width=20, command=lambda: self.getPassword()).grid(row=4,column=0,columnspan=3,pady=30)

    def getPassword(self):
        self.controller.username = self.email_entry.get()
        self.controller.password = self.password_entry.get()
        #print(self.controller.password)
        self.controller.connectToServer()


class PageOne(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        #connection = controller.connectToServer(parent.password)

        tk.Label(self, text="Total Emails in Inbox: ").grid(row=0,sticky='W',padx=30,pady=[10,0], columnspan=1)
        tk.Label(self, textvariable=self.controller.numOfMessages).grid(row=0,sticky='W',column=1,pady=[10,0], columnspan=1)

        self.variable = tk.StringVar(self)
        self.variable.set("INBOX") #defaultvalue


        tk.Label(self, text="Choose Mailbox: ").grid(row=1,sticky='W',padx=30,pady=10, columnspan=1) #temporarily suspended
        self.mailboxes = tk.OptionMenu(self,self.variable,"INBOX","Dive Team","Lap Swim", "Swim Team", "Water Polo", "History")
        #self.mailboxes = tk.OptionMenu(self, self.variable, *self.controller.list)
        #self.mailboxes = tk.OptionMenu(self, variable, tuple(self.controller.list))
        #self.mailboxes = apply(OptionMenu, (self, variable) + tuple(self.controller.list))

        #self.mailboxes.grid(row=1,column=1,sticky='EW',pady=10, columnspan=1) #temporarily suspended
        self.mailboxes.config(width=10)

        label = tk.Label(self, text="-----------------------------", font=controller.title_font).grid(row=1,sticky='EW',padx=30,pady=0, columnspan=2)

        tk.Label(self, text="Number of Records: ").grid(row=2,sticky='W',padx=30,pady=10, columnspan=1)
        self.number_of_waivers = tk.Entry(self, width=5)
        self.number_of_waivers.insert(0,"0")
        self.number_of_waivers.grid(row=2,sticky='W', pady=10, column=1)
        self.error = tk.Label(self, text="Input must be less than total emails.", fg="red")
        #label = tk.Label(self, text="----- Or -----", font=controller.title_font).grid(row=3,sticky='EW',padx=30,pady=10, columnspan=2)

        min = "6/24/2020"
        min = min.split('/')
        min = datetime.datetime(int(min[2]), int(min[0]), int(min[1]) ).date()
        today = date.today()

        #tk.Label(self, text="Start Date: ").grid(row=4,sticky='W',padx=30,pady=10, columnspan=1)
        self.startDate = DateEntry(self,minDate=min,maxDate=today)
        #self.startDate.grid(row=4,sticky='W',column=1,pady=10)

        #tk.Label(self, text="End Date:").grid(row=5,sticky='W',padx=30,pady=10, columnspan=1)
        self.endDate = DateEntry(self,minDate=min,maxDate=today)
        #self.endDate.grid(row=5,sticky='W',column=1,pady=10)
        #https://tkcalendar.readthedocs.io/en/stable/_modules/tkcalendar/calendar_.html

        login_button = tk.Button(self,text="Submit", width=20, command=lambda: self.getData()).grid(row=7,column=0,columnspan=2,pady=10)

    def getData(self):
        #self.show_frame("PageTwo")
        numOfWaivers = int(self.number_of_waivers.get())
        startDate = self.startDate.get_date()
        endDate = self.endDate.get_date()

        if(numOfWaivers > self.controller.intOfMessages):
            self.error.grid(row=3,sticky='EW',padx=30,pady=10, columnspan=2)
        else:
            self.error.grid_forget()
            self.controller.show_frame("PageTwo")aaaaa  BV ,M                     
            self.controller.get_waivers_by_number(numOfWaivers,self.variable.get())
            self.controller.show_frame("PageThree")

class PageTwo(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text="Processing Request...\nPlease Wait", font=controller.title_font).pack(pady=100)
        #exit_button = tk.Button(self,text="OK", width=20, command=lambda: exit()).pack()

class PageThree(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text="Request Successful.", font=controller.title_font).pack(pady=30)
        label2 = tk.Label(self, text="Please turn off permissions for less secure apps.").pack()
        label3 = tk.Label(self, text="Move all remaining emails in inbox to archive.").pack(pady=20)
        exit_button = tk.Button(self,text="OK", width=20, command=lambda: exit()).pack()

if __name__ == "__main__":
    app = WaiverApp()
    app.mainloop()
