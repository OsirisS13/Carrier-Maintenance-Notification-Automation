# -*- coding: utf-8 -*-
## TO DO:
## Change regex so that multiple circuit IDs are matched, if they are present https://stackoverflow.com/questions/20240239/python-re-search DONE for euNetworks
## Investigate why schedule conflict detection is not working properly, sometimes it returns results outside the period, other times it does not return a result within the period
## Move email to an archive folder once event has been added to calendar
## Package into an installer http://www.pyinstaller.org/ http://takluyver.github.io/posts/so-you-want-to-write-a-desktop-app-in-python.html

#regex module
import re
import datetime
#databse module
import pyodbc 
#interact with windows programs
import win32com.client

#Add the mailbox to search for maintenance ID's here
mailbox = ""
#Address or mailing group to send outlook invite to
appointment_recipient = ""

# Read emails
class OutlookLib:
#from http://softwaretestautomationnotes.blogspot.nl/2011/11/reading-outlook-email-with-python.html         
    def __init__(self, settings={}):
        self.settings = settings
        # function to read inbox
    def get_messages(self, user, folder="Inbox", match_field="all", match="all"):      
        outlook = win32com.client.Dispatch("Outlook.Application")
        myfolder = outlook.GetNamespace("MAPI").Folders[user] 
        inbox = myfolder.Folders[folder] # Inbox
        if match_field == "all" and match =="all":
            return inbox.Items
        else:
            messages = []
            for msg in inbox.Items:
                try:
                    if match_field == "Sender":
                        if msg.SenderName.find(match) >= 0:
                            messages.append(msg)
                    elif match_field == "Subject":
                        if msg.Subject.find(match) >= 0:
                            messages.append(msg)
                    elif match_field == "Body":
                        if msg.Body.find(match) >= 0:
                            messages.append(msg)
                except:
                    pass
            return messages
         
    def get_body(self, msg):
        return msg.HTMLBody
     
    def get_subject(self, msg):
        return msg.Subject
     
    def get_sender(self, msg):
        return msg.SenderName
     
    def get_recipient(self, msg):
        return msg.To
     
    def get_attachments(self, msg):
        return msg.Attachments
         
		 
def eunetworks_maintenance():
	body = msg.Body.encode("utf-8") 
	#regex here matches everything after the ?<= (inside the brackets) until the end of the line, and then strips any carriage returns
	start_time = re.search('(?<=Start Time: ).*', body).group(0).rstrip()
	#convert start time from string to datetime object, see https://docs.python.org/2/library/datetime.html#strftime-strptime-behavior
	start_time =  datetime.datetime.strptime(start_time, '%Y-%m-%d %H:%M %Z')
	end_time = re.search('(?<=End Time: ).*', body).group(0).rstrip()
	#convert end time from string to datetime object, see url above
	end_time =  datetime.datetime.strptime(end_time, '%Y-%m-%d %H:%M %Z')
	#regex captures everything after a comma and space, that has a C as the next character, until the next comma, 
	#but excludes Cessnalaan (as some service have their A or Z points here)
	circuit_id = re.findall('(?!Cessnalaan)(?<=, )(C[^,]*)', body, re.MULTILINE)
	subject = msg.Subject
	#return multiple values as a dictionary
	return {"circuit_id": circuit_id, "start_time": start_time, "end_time": end_time, "subject": subject, "body": body}
	
	
def interoute_maintenance():
	body = msg.Body.encode('utf-8') 
	#regex here matches everything after the ?<= (inside the brackets) until the the next occurence of a open bracket "(" , and then strips any carriage returns from the result
	#try the first regex for date and time, if that returns none try the next (as the text is different for new maintenances vs update notifications where the date has been changed
	start_time = re.search('(?<=Start of Planned Work Window: )[^(]*', body)#.group(0).rstrip()
	if start_time == None:
		start_time = re.search('(?<=New Planned Work Start Date: )[^(]*', body).group(0).rstrip()
	else:
		start_time = start_time.group(0).rstrip()
	#convert start time from string to datetime object, see https://docs.python.org/2/library/datetime.html#strftime-strptime-behavior
	start_time =  datetime.datetime.strptime(start_time, '%d/%b/%Y %H:%M %Z')
	end_time = re.search('(?<=End of Planned Work Window: )[^(]*', body)#.group(0).rstrip()
	if end_time == None:
		end_time = re.search('(?<=New Planned Work End Date: )[^(]*', body).group(0).rstrip()
	else:
		end_time = end_time.group(0).rstrip()
	#convert end time from string to datetime object, see url above
	end_time =  datetime.datetime.strptime(end_time, '%d/%b/%Y %H:%M %Z')
	#for this we need to replace any new line \n or \r in order to be able to parse with regex
	body_stripped_of_newlines =  body.replace('\n', '')
	body_stripped_of_newlines =  body_stripped_of_newlines.replace('\r', '')
	#then run regex for the string inside the brackets (inc spaces) until the next space to get the circuit id
	circuit_id = re.search('(?<=Friendly Name:  )[^ ]*', body_stripped_of_newlines).group(0).rstrip()
	subject = msg.Subject
	return {"circuit_id": circuit_id, "start_time": start_time, "end_time": end_time, "subject": subject, "body": body}
	
	
def telia_maintenance():
	body = msg.Body.encode('utf-8') 
	#regex here matches everything after the ?<= (inside the brackets) until the end of the line, and then strips any carriage returns from the result
	start_time = re.search('(?<=Start Date and Time: ).*', body).group(0).rstrip()
	#convert start time from string to datetime object, see https://docs.python.org/2/library/datetime.html#strftime-strptime-behavior
	start_time =  datetime.datetime.strptime(start_time, '%Y-%b-%d %H:%M %Z')
	print start_time
	end_time = re.search('(?<=End Date and Time: ).*[^(]*', body).group(0).rstrip()
	#convert end time from string to datetime object, see url above
	end_time =  datetime.datetime.strptime(end_time, '%Y-%b-%d %H:%M %Z')
	print end_time
	#regex here matches everything after the ?<= (inside the brackets) until the end of the line, and then strips any carriage returns from the result
	circuit_id = re.search('(?<=Service ID: ).*', body).group(0).rstrip()
	subject = msg.Subject
	return {"circuit_id": circuit_id, "start_time": start_time, "end_time": end_time, "subject": subject, "body": body}
	
	
def level3_maintenance(maintenance_id):
	body = msg.Body.encode('utf-8') 
	subject = msg.Subject
	#regex here matches everything after the ?<= (inside the brackets) until the end of the line, and then strips any carriage returns from the result
	#need to put regex in a string as we are using a variable inside it, and cant just build the regex during the search
	#-1 at the end to match the maint ID with -1 as per email format
	start_regex_search = '(?<=' + maintenance_id + '-1 )[^to]*'
	start_time = re.search(start_regex_search, body).group(0).rstrip()
	#build regex string using the start time as the variable, until end of line
	end_regex_search = '(?<=' + start_time + ' to ).*'
	end_time = re.search(end_regex_search, body).group(0).rstrip()
	#convert start time from string to datetime object, see https://docs.python.org/2/library/datetime.html#strftime-strptime-behavior
	start_time =  datetime.datetime.strptime(start_time, '%Y-%m-%d %H:%M:%S %Z')
	#convert end time from string to datetime object, see url above
	end_time =  datetime.datetime.strptime(end_time, '%Y-%m-%d %H:%M:%S %Z')
	#regex here matches everything after the ?<= (inside the brackets inc spaces) until the next </td>, and then strips any carriage returns from the result
	#using html body of email.  Need to fill out your customer name here
	customer_name = ""
	circuit_regex_search = '(?<=<td>' + customer_name + '</td><td>)[^</td>]*'
	circuit_id = re.findall(circuit_regex_search, msg.HTMLbody, re.MULTILINE)
	return{"circuit_id": circuit_id, "start_time": start_time, "end_time": end_time, "subject": subject, "body": body}
	
	
#database lookup, see https://docs.microsoft.com/en-us/sql/connect/python/pyodbc/step-3-proof-of-concept-connecting-to-sql-using-pyodbc?view=sql-server-2017
#this function is specific to the database you keep your circuit information in.  SQL query here is generic and will need to be modified for your DB
def lookup_circuitID(circuit_id):
	server = "" 
	database = ""
	username = "" 
	password = "" 
	#create connection to DB using values specified above using pydobc module, and the SQL Server driver
	cnxn = pyodbc.connect("DRIVER={SQL Server};SERVER="+server+";DATABASE="+database+";UID="+username+";PWD="+ password)
	cursor = cnxn.cursor()
	#build and the execute the below query against the DB specified for the circuit id
	sql_query ="SELECT * FROM CircuitTable WHERE CircuitID LIKE '" + circuit_id + "%';"
	cursor.execute(sql_query) 
	row = cursor.fetchone() 
	try:
		circuit_description = row.Description
		circuit_purpose = row.Purpose
		return {"circuit_description": circuit_description, "circuit_purpose": circuit_purpose}
	except:
		print "Error!  Could not get Circuit information from database!"
		exit()
		return
	
		
#function to create the calendar event
#http://www.baryudin.com/blog/sending-outlook-appointments-python.html
def add_calendar_event(circuit_id, circuit_description, circuit_purpose, start_time, end_time, subject, body):
	Outlook = win32com.client.Dispatch("Outlook.Application")
	appointment = Outlook.CreateItem(1) # "1" = outlook appointment item https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/olitemtype-enumeration-outlook
	#turn appointment into a meeting
	appointment.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. only after changing the meeting status recipients can be added https://msdn.microsoft.com/en-us/VBA/Outlook-VBA/articles/olmeetingstatus-enumeration-outlook
	appointment.Start = start_time
	appointment.End = end_time
	appointment.Subject = circuit_description + " " + circuit_purpose + " " + subject
	appointment.Body = body
	appointment.ReminderSet = True
	appointment.ReminderMinutesBeforeStart = 15
	appointment.Recipients.Add(appointment_recipient) 
	appointment.Save()
	appointment.Send()
	print "Created event: " + circuit_description + " " + circuit_purpose + " " + subject
	return
	
	
def check_conflicting_events(start_time, end_time):
#init and open the outlook reader
	outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
	calendar = outlook.GetDefaultFolder(9)  #"9" refers to the index of a folder - in this case, the events/appointments.  https://stackoverflow.com/questions/38899956/python-win32com-get-outlook-event-appointment-meeting-response-status
	#supposed to exclude recurring events but didnt seem to work 
	#calendar.Items.IncludeRecurrences = False
	#assign all returned items into variable
	appointments = calendar.Items
	#to match all events during this period, check for existing events that start before the end of the new maintenance, and existing events that end after the start time of the new maintenance
	check_times = "[Start] <= '" + end_time.strftime("%m/%d/%Y %I:%M%p") + "' AND [End] >= '" + start_time.strftime("%m/%d/%Y %I:%M%p") + "'" #for explanation check https://stackoverflow.com/questions/21477599/read-outlook-events-via-python
	#after applying filter, assign returned items into variable
	restrictedItems = appointments.Restrict(check_times)
	print "\nChecking for existing events between", start_time.strftime("%m/%d/%Y %I:%M%p"), "and",end_time.strftime("%m/%d/%Y %I:%M%p")
	#if a count of the returned items is 0, notify user
	if restrictedItems.Count == 0:
		print "No other maintenances during this time!"
	#otherwise notify user how many outlook events occur at the same time, and list them all
	else:	
		print "Found", restrictedItems.Count,"other events during this time, please check to make sure they are not maintenances that clash: \n ---------------------------------------------"
		for item in restrictedItems:
				print item.Subject
				print "Starting:", item.Start
				print "Ending:", item.End
				print "---------------------------------------------"
		
#user input to begin script
maintenance_to_search_for = raw_input("Enter Maintenance ID: ")
outlook = OutlookLib()
#mailbox to check messages for
messages = outlook.get_messages(mailbox)
#loop through all messages in mailbox
for msg in messages:
	#until matching the inputted string in the subject of an email
	if maintenance_to_search_for in msg.Subject:
		#get the email address of the sender
		sender = msg.SenderEmailAddress
		#and kickoff relevant function to parse the message and extract the needed info, which then returns them formatted nicely in a dictionary
		if sender == "change.management.EMEA@Level3.com":
			maintenance_values = level3_maintenance(maintenance_to_search_for)
		elif sender == "maintenance@eunetworks.com":
			maintenance_values = eunetworks_maintenance()
		elif sender == "netopsadmin@interoute.com":
			maintenance_values = interoute_maintenance()
		elif sender == "ncm@teliacompany.com":
			maintenance_values = telia_maintenance()
		#move email to appropriate folder, this does not work (probably incorrect folder structure)
		# msg.Move(win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").Folders[mailbox].Folders["Inbox"].Folders["Archive"])
		     


#Values needed to create the maintenance calendar event are stored in the dictionary.
#Here we access a single value in the dictonary and assign it to the relevant variable
try:
	circuit_id = maintenance_values["circuit_id"]
	start_time = maintenance_values["start_time"]
	end_time = maintenance_values["end_time"]
	subject =  maintenance_values["subject"]
	body =  maintenance_values["body"]
	print circuit_id
except NameError:
	print "ERROR! Could not find an email with a Maintenance ID matching", maintenance_to_search_for
	exit()
#init list where the returned circuit descriptions are stored
descriptions = ''
#if circuit id returns as a list (ie more than one result, run lookup circuit function using the circuit id as a loop,  otherwise just assign normally)
if isinstance(circuit_id, list):
	all_circuits = []
	for circuit in circuit_id:
		circuit_details = lookup_circuitID(circuit)
		all_circuits.append(circuit_details)
		circuit_purpose = circuit_details["circuit_purpose"]
		#loop through if list length is greater than 1
	if len(circuit_id) > 1:
		for circuit in all_circuits:
			descriptions = str(descriptions + circuit["circuit_description"]) +  ' | '
	else:
		descriptions = circuit_details["circuit_description"]
else: 
	circuit_details = lookup_circuitID(circuit_id)
	descriptions = circuit_details["circuit_description"]
	circuit_purpose = circuit_details["circuit_purpose"]

#run conflict checker function
check_conflicting_events(start_time, end_time)
#run function to add event to outlook calendar
add_calendar_event(circuit_id, descriptions, circuit_purpose, start_time, end_time, subject, body)



