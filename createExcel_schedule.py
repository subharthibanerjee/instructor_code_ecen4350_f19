"""
	__author__="Subharthi"
	__date__=""
	__version__="v1.0.4"

	creates an excel file where it creates entries for
	class dates 

	Requirements: install beautifulsoup, xlsxwriter
"""
from __future__ import print_function
import logging
import xlsxwriter 
  

from sys import platform
import os, os.path
from datetime import date, timedelta

from datetime import datetime

from collections import namedtuple


import time
import urllib.request 

from bs4 import BeautifulSoup

import re


import datetime
import pickle

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

__author__="Subharthi"
__date__="08/17/2018"
__version__="v1.0.3"


# ===================================================================
format = "%(asctime)s: %(message)s"
logging.basicConfig(format=format, level=logging.INFO,
                        datefmt="%H:%M:%S") # logging information
# ===================================================================

xl_filename = "ecen4350f19_sched.xlsx"

# ===================================================================
# column description
n_topics=7
task_cols = ["Task "+str(i) for i in range(1, n_topics+1)]

labs_hw = ['Lab Today', 'Lab Assignment Posted', 'Lab Assignment Due', 
			'HW Assignment Posted', 'HW Assignment Due',
			'Special Notice']
date_header = "Class Dates"

task_cols.insert(0, "Lecture No.")
task_cols.insert(0,date_header)

for data in labs_hw:
	task_cols.append(data)
task_cols_n = [i for i in range(0, len(task_cols))]
task_cols_dict = dict(zip(task_cols, task_cols_n))

# ====================== calendar dates ============================

# this is for calendar events to work with google calendar ---
calender_events = None
SCOPES = 'https://www.googleapis.com/auth/calendar'


allEventsForCalendar = []

all_generated_files = []
# ===================================================================






if_quiz = False
if_meeting = False
if_holiday = False
if_exam=False



lab_description=["Implement Serial", "First Meeting","SPI Lab", "I2C lab",
				"Schematic Design", "PCB design"]
n_labs=len(lab_description)
lab_numbers = ["Lab Assignment - "+str(i) for i in range(1, n_labs+1)]
hw_description=["Initial Block Diagram", "First Meeting", "Final Block Diagram", 
				"Component Selection"]
n_hws = len(hw_description)

hw_numbers = ["HW Assignment - "+str(i) for i in range(1, n_labs+1)]


lab_dict=dict(zip(lab_numbers, lab_description))
hw_dict=dict(zip(hw_numbers, hw_description))


lab_post_date = None

exam_post = None
lab_due_dates = None

project_demo_due = None


lab_post_dates=[i for i in range(1, n_labs+1)]
lab_due_dates=[i+1 for i in range(1, n_labs+1)]

hw_post_dates=[i for i in range(1, n_hws+1)]
hw_due_dates=[i+1 for i in range(1, n_hws+1)]




pcb_due = 9

project_demo_due=15
project_report_due=16

exam_1 = 4
exam_2 = 8
final_exam = 16

# ===================================================================
year=2019
fall_month=8
fall_day=26

fall_end_day=12
fall_end_month=12

MONTHS_TO_CHECK = [i for i in range(fall_month+1, fall_end_month+1)]
#print(MONTHS_TO_CHECK)
MONTHS_TO_CHECK_NAME = [date(1900, i, 1).strftime('%B') for i in MONTHS_TO_CHECK]


# ===========================================================
"""
writing code to extract holidays from uno_site
"""
uno_site = "https://www.unomaha.edu/registrar/academic-calendar.php"
unl_site = "https://registrar.unl.edu/academic-calendar/archive/2019-2020/"

unmc_site = "http://catalog.unmc.edu/general-information/academic-calendar2/"

test_site = "https://bgcmalibu.org/vc_block/club-holiday-schedule/"

temp_file = "test_soup.txt"
uno_temp_file = "uno_test_soup.txt"
unl_temp_file = "unl_test_soup.txt"
unmc_temp_file = "unmc_test_soup.txt"
test_temp_file = 'test_test_soup.txt'


event_file_name = "eventIds.txt"

all_generated_files.append(temp_file)
all_generated_files.append(uno_temp_file)
all_generated_files.append(unl_temp_file)
all_generated_files.append(unmc_temp_file)
all_generated_files.append(test_temp_file)
all_generated_files.append(xl_filename)
all_generated_files.append('token.pickle')
all_generated_files.append(xl_filename)

# TODO: check for unmc site, there is a nonetype

def findHolidays(site, tmp_file):
	"""
	rtype:
	"""

	if os.path.isfile(tmp_file):
		logging.info("[IN findHolidays] File already exists, opening from %s", tmp_file)
		if os.stat(tmp_file).st_size==0:
			logging.info("[IN findHolidays] File %s already exists, but size=0", tmp_file)
			os.remove(tmp_file)
			findHolidays(site)
		else:
			with open(tmp_file,"r", encoding="utf8") as f:
				content = f.read()
			return content
	else:
		logging.info("File does not exist, opening from internet %s", site)
		if platform=='darwin':
			logging.info('Unsecure ssl verification off http request')
			import ssl
			if (not os.environ.get('PYTHONHTTPSVERIFY', '') and getattr(ssl, '_create_unverified_context', None)): 
				ssl._create_default_https_context = ssl._create_unverified_context
		
		url = urllib.request.urlopen(site)

		content = url.read()
		if len(content) is 0:
			logging.info("File creation failed due to length of content scraped is %d", len(content))
		else:
			with open(tmp_file, "wb") as f:
				f.write(content)
			logging.info("[IN findHolidays] File write completed, %s created ", tmp_file)
		return content


	
# ======================================


def appendTextHolidays(reason, tmp_file):

	"""
	rtype:
	"""
	filename = tmp_file.replace('.txt','')
	filename = filename + '_holiday_with_reason' +'.txt'
	if os.path.isfile(filename):
		logging.info("[IN appendTextHolidays] File already exists, opening from %s", filename)
		if os.stat(filename).st_size==0:
			logging.info("[IN appendTextHolidays] File %s already exists, but size=0", filename)
			os.remove(filename)
			appendTextHolidays(reason, tmp_file)

	else:
		logging.info('File %s created and writing', filename)
		if len(reason) is 0:
			logging.info("[FILE] %s is not created, not sufficient length", filename)
		else:
			with open(filename, 'w', encoding="utf8") as f:
				f.write('\n'.join(reason))
			logging.info('File %s created and written', filename)
	return filename


# ======================================

def checkHolidays(site, tmp_file):
	"""
	check holidays: find holidays in the text file, scrape those rows, find parent 
	element and list the dates. Use soup

	rtype: hashset of strings
	"""
	holidays = []
	holidays_with_reason = []
	content = findHolidays(site, tmp_file)
	soup = BeautifulSoup(content, 
		features="html.parser")
	soup.prettify()
	print("\n\n\n")
	logging.info("Checking for holidays from academic-calendar %s\n\n", site)
	for tag in soup.find_all(text=re.compile('Holiday')):
		if site is unmc_site:
			col = tag.findParent('tbody')
		if site is test_site:
			col = tag.findParent('ul')

		if site is uno_site:
			col = tag.findParent('tr')
		else:
			col = tag.findParent('tr')
		#print("tag: ",col)

		
		
		if col is not None:
			cols = col.get_text() 
			newcol = cols.replace('\n', ' ')
			holiday_with_reason=newcol
			holidays_with_reason.append(holiday_with_reason)
		

			# check for month names in the semester
			# that matches with holidays
			
			
			#print(holiday_with_reason)
			for month in MONTHS_TO_CHECK_NAME:
				
				if len(col.find_all(text=re.compile(month))) is not 0:
					holiday = col.find_all(text=re.compile(month))
					
					#print(holiday)
					holidays.extend(holiday)

	fname = appendTextHolidays(holidays_with_reason, tmp_file)
	all_generated_files.append(fname)
	print("file name created with holidays ",fname)

	#print("reason: ", holidays_with_reason)
	return list(set(holidays)) # use hashset for no duplicates
		
# ============================================================






def extracttMonthDay(literal, year):
	"""
	rtype:

	"""

	months = [i for i in range(1, 13)]
	#print(MONTHS_TO_CHECK)
	month_names = [date(1900, i, 1).strftime('%B') for i in months]
	# make dictionary

	hl_start_m, hl_start_d, hl_end_m, hl_end_d = 0, 0, 0, 0
	MONTH_DICT = dict(zip(month_names, months))
	days = [str(i) for i in range(1, 32)]
	#print(days)
	#month = re.findall(r"\w+",literal)
	month = re.findall(r"\w+",literal)
	#print(month)
	months_n = []
	days_n = []
	# strip all the unnecessary values.
	# we can do years too here
	for ele in month:
		
		if (ele in month_names):
			
			months_n.append(ele)
		elif ele in days:
			days_n.append(ele)

	
	if len(months_n) == 2:
		hl_start_m = months_n[0]
		hl_end_m = months_n[1]
	elif len(months_n) == 1:
		hl_start_m = months_n[0]
		hl_end_m = hl_start_m

	if len(days_n) == 2:
		hl_start_d = days_n[0]
		hl_end_d = days_n[1]
	elif len(days_n) == 1:
		hl_start_d = days_n[0]
		hl_end_d = hl_start_d

	
	
	hl_start = date(year, MONTH_DICT[hl_start_m], int(hl_start_d))
	hl_end = date(year, MONTH_DICT[hl_end_m], int(hl_end_d))
			
			
		
			

	return [hl_start, hl_end]

# ==========================================+++++++++++++++++++++

def extractDateTimes(holiday_strings, year, ifstring):
	"""
	rtype:
	"""
	sem_holidays = []
	sem_holidays_dict=None
	for d in holiday_strings:
		if not ifstring:
			sem_holidays.append(extracttMonthDay(d, year))
		else:
			dates = extracttMonthDay(d, year)
			sem_holidays.append(dates[0]) 
			alldays = dates[1] - dates[0]

			for allday in range(0, alldays.days):
				holidates = dates[0]+timedelta(days=allday)
				
				sem_holidays.append(holidates) 
			#for date in dates:
			#	sem_holidays.append(date) 
	return sem_holidays

def createSemStartEnd(sem_start, sem_end):
	"""

	rtype: List[namedtuple]
	"""
	semester_start_date=date(sem_start.year, 
						sem_start.month, sem_start.day)

	semester_end_date=date(sem_end.year, 
						sem_end.month, sem_end.day)


	return [semester_start_date, semester_end_date]
	

def listAllSemesterdays(sem_start, sem_end):
	"""
	rtype:
	"""

	allTuesdays = []
	tu_dict=dict()
	th_dict=dict()
	allThursdays = []
	sem_start_date, sem_end_date = createSemStartEnd(sem_start, sem_end)
	days_in_semester = sem_end_date - sem_start_date
	print("there are ", days_in_semester.days, " days in this semester")
	for i in range(1, days_in_semester.days + 1, 7):
		tuesday = sem_start_date + timedelta(days=i)
		thursday = sem_start_date + timedelta(days=i+2)
		tu_dict[tuesday]="Tue"
		th_dict[thursday]="Thurs"
		allTuesdays.append(tuesday)
		allThursdays.append(thursday)

	allsemDays= [None]*(len(allTuesdays)+len(allThursdays))
	allsemDays[::2]=allTuesdays
	allsemDays[1::2]=allThursdays
	return allsemDays, tu_dict, th_dict


####========================================================

def printDateTime(datetime_obj):
	"""
	rtype: 
	"""
	print()
	for dt in datetime_obj:
		if not isinstance(dt, date):
			for d in dt:
				print(d, end=", ")
                               
		else:
			print(dt, end=", ")
                       

		print()


def createFormatExcel(formatting):
	"""
	"""

	
	tags = ['bold', 'bg_color', 
			'font_color', 'font_size', 'text_wrap']
	format_dict=dict(zip(tags, formatting))
	return format_dict
	




def createListofEvents(summary, description, week_days):
	"""
	"""

	startTime = {
	'dateTime':datetime.datetime.now().isoformat(), 
	'timeZone':'America/Chicago'
	}
	endTime = {
		'dateTime':datetime.datetime.now().isoformat(), 
		'timeZone':'America/Chicago'
	}

	attendee = {'email':None}

	reminder= {
		'useDefault':True,
		'overrides':[]	
	} 


	room_location = 'PKI-256'

	default_event = {
		'summary': None, 
		'location':room_location,
		'description': None, 
		'start':startTime, 
		'end':endTime, 
		'recurrence':[],
		'attendees':[],
		'reminders': reminder,
	}

	
	default_event['summary'] = summary
	default_event['description'] = description
	if len(week_days) < 1:
		pass
	elif len(week_days) < 2:
		week_days[0]=datetime.datetime.combine(week_days[0], datetime.datetime.min.time())
		

		startTime['dateTime'] = week_days[0].isoformat()

		endTime['dateTime'] = week_days[0].isoformat()
		

	else:
		week_days[0]=datetime.datetime.combine(week_days[0], datetime.datetime.min.time())
		week_days[1]=datetime.datetime.combine(week_days[1], datetime.datetime.min.time())
		startTime['dateTime'] = week_days[0].isoformat()

		endTime['dateTime'] = week_days[1].isoformat()

	
	default_event['start'] = startTime
	default_event['end'] = endTime

	print(default_event)
	
	return default_event


def createExcel(header, writeables, filename, holidays, ifevent=True):
	"""
	rtype:
	"""
	row, col = 0, 0
	workbook = xlsxwriter.Workbook(filename) 
	worksheet = workbook.add_worksheet("sched1")
	logging.info("opened workbook %s for worksheet", filename)


	format_cell = [True, 'cyan',  
				'black', '14', True]
	format_dict=createFormatExcel(format_cell)
	cell_format = workbook.add_format(format_dict)
	worksheet.set_row(0, 70)
	
	for head in header:
		
		worksheet.write(row, col, head,  cell_format)
		col = col+1
	row = row + 1
	col = 0
	count = 2
	week = 0
	
	# TODO Please rewrite this section to make the code more interpretable

	# each week startdate and end date with event gets added 
	# to calendat if ifEvent is True
	week_days = []
	list_of_events= []
	for ele in writeables:

		week_days.append(ele)

		if count == 2:
			
			week = week + 1
			Week = "Week - " + str(week)
			hw_numbers_iter = iter(hw_numbers)
			lab_numbers_iter = iter(lab_numbers)
			format_cell = [True, 'white', 
				'black', '8', True]
			format_dict=createFormatExcel(format_cell)
			cell_format = workbook.add_format(format_dict)


			# @TODO: make a function for posting hw
			if week in hw_post_dates:

				text = hw_numbers[week-1]+'('+ hw_dict[hw_numbers[week-1]]+')'
				
				worksheet.write(row, task_cols_dict['HW Assignment Posted'],
								 text, cell_format)
				
				# if event make a disctionary to post in calendar
				if ifevent:
					list_of_events.append(createListofEvents(hw_numbers[week - 1], text, [ele, ele+timedelta(7)]))
					


			if week in lab_post_dates:
				text = lab_numbers[week-1]+'('+lab_dict[lab_numbers[week-1]] +')'
				worksheet.write(row, task_cols_dict['Lab Assignment Posted'],
								 text, cell_format)
				if ifevent:
					list_of_events.append(createListofEvents(lab_numbers[week - 1], text, [ele, ele+timedelta(7)]))


			if week in lab_due_dates:
				text = lab_numbers[week-2]+'('+ lab_dict[lab_numbers[week-2]]+')'
								 
				worksheet.write(row, task_cols_dict['Lab Assignment Due'],
								 text
								 , cell_format)
				#if ifevent:
				#	list_of_events.append(createListofEvents(defaultEvent, 
				#		lab_numbers[week - 2], text, week_days))


			if week in hw_due_dates:
				text =  hw_numbers[week-2]+'('+hw_dict[hw_numbers[week-2]]+')'
				worksheet.write(row, task_cols_dict['HW Assignment Due'],
								text, cell_format)

				#if ifevent:
				#	list_of_events.append(createListofEvents(defaultEvent, 
				#		hw_numbers[week - 2], text, week_days))


			if week == project_demo_due:
				worksheet.write(row, task_cols_dict['Special Notice'],
								'Project Demo Due', cell_format)

				if ifevent:
					list_of_events.append(createListofEvents('Project Demo Due', 'Project Demo Due', [ele, ele+timedelta(7)]))

			if week == project_report_due:
				worksheet.write(row, task_cols_dict['Special Notice'],
								'Project Report Due', cell_format)
				if ifevent:
					list_of_events.append(createListofEvents('Project Report Due', 'Project Report Due', [ele, ele+timedelta(7)]))

			if week == pcb_due:
				worksheet.write(row, task_cols_dict['Special Notice'],
								'PCB Submission to Manufacturing Due', cell_format)

				if ifevent:
					list_of_events.append(createListofEvents('PCB Submission to Manufacturing Due', 'PCB Submission to Manufacturing Due', [ele, ele+timedelta(7)]))



			# @TODO: make a function for posting exam 
			if week == exam_1:
			
				for c in range(len(header[:task_cols_dict['Lab Today']])):
					format_cell = [True, 'green', 
									'black', '8', True]
					format_dict=createFormatExcel(format_cell)
					cell_format = workbook.add_format(format_dict)

					worksheet.write(row, c,
								'Exam 1', cell_format)
				worksheet.write(row, task_cols_dict['Special Notice'],
								'Exam 1 today', cell_format)

				if ifevent:
					list_of_events.append(createListofEvents('Exam 1', 'First Exam', [ele]))




			if week == exam_2:
				for c in range(len(header[:task_cols_dict['Lab Today']])):
					format_cell = [True, 'green', 
									'black', '8', True]
					format_dict=createFormatExcel(format_cell)
					cell_format = workbook.add_format(format_dict)

					worksheet.write(row, c,
								'Exam 2', cell_format)
				worksheet.write(row, task_cols_dict['Special Notice'],
								'Exam 2 today', cell_format)


				if ifevent:
					list_of_events.append(createListofEvents('Exam 2', 'Second Exam', [ele]))



			if week == final_exam:
				# if there is an exam please make all the cells green
				for c in range(len(header[:task_cols_dict['Lab Today']])):
					format_cell = [True, 'green', 
									'black', '8', True]
					format_dict=createFormatExcel(format_cell)
					cell_format = workbook.add_format(format_dict)

					worksheet.write(row, c,
								'Final Exam', cell_format)
				worksheet.write(row, task_cols_dict['Special Notice'],
								'Final Exam Today', cell_format)



				if ifevent:
					list_of_events.append(createListofEvents('Final Exam', 'Final Exam', [ele]))
	
				
			logging.info("Writing %s", Week)
			format_cell = [True, 'yellow', 
				'black', '10', True]
			format_dict=createFormatExcel(format_cell)
			cell_format = workbook.add_format(format_dict)
			worksheet.write(row, col, Week, cell_format)
			row = row + 1
			count = 0
			week_days = []
		#worksheet.write(row, col, column)
		if ele in holidays:
			# if holidays make column red
			
			format_cell = [True, 'red', 
				'black', '10', True]
			format_dict=createFormatExcel(format_cell)
			cell_format = workbook.add_format(format_dict)
			ele = ele.strftime("%m/%d/%Y")
			worksheet.write(row, 0, ele, cell_format)
			logging.info("Holidays: no class")
			for cell in range(1, len(header)):
				worksheet.write(row, cell, "No Class", cell_format)
			row = row + 1
			count = count + 1
			
			
		else:
			
			ele = ele.strftime("%m/%d/%Y")
			worksheet.write(row, col, ele)
			logging.info("writing %s ", ele)
			row = row + 1
			count = count + 1
			


	logging.info("finished creating workbook")
	workbook.close()
	
	#print("allEventsForCalendar ", allEventsForCalendar)
	return list_of_events



def createAssignment(posted_dates, due_in, lab_assignment):
	"""
	rtype:
	"""
	pass






def createEventGoogleCalendar(eventList, eventfile):
	""" Create events from the stored events due to the creation
	of Excel sheet
	rtype:
	"""

	creds = None	

	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)

	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				'credentials.json', SCOPES)
			
			creds = flow.run_local_server(port=0)

			# save the credentials for next run
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)
			
	print("service creds: ", creds)
	service = build('calendar', 'v3', credentials=creds, cache_discovery=False)

	# call the calendar API

	#now = datetime.datetime.utcnow().isoformat()+'Z' # Z indicates UTC time
	#print(now)
	logging.info('Print Events in the calendar ')

	#count = 0
	eventId = []
	for event in eventList:

		
		#eventId = 'event' + str(count)
		
		logging.info('for loop ')
		event = service.events().insert(calendarId='primary', body=event).execute()
		eventId.append(event['id'])
		print('event id:', eventId)
		logging.info('Event created: %s', (event.get('htmlLink')))
		#count = count + 1


	if os.path.isfile(eventfile):
		logging.info("Event id file exists.")
		os.remove(eventfile)
		logging.info("removing previous id files")
	else:
		with open(eventfile, "wb") as f:
			# use when np picle available
			#f.writelines("%s\n", eventid for eventid in eventId)
			pickle.dump(eventId, f)







def cleanUp(eventfilename):
	"""
	"""


	for filename in all_generated_files:
		
		if os.path.isfile(filename):
			logging.info("file %s exists, removing", filename)
			os.remove(filename)
			print("removed file {}".format(filename))

	logging.info("Removing all previously generated calendar events")

	creds = None	

	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)

	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				'credentials.json', SCOPES)
			
			creds = flow.run_local_server(port=0)

			# save the credentials for next run
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)
			
	print("service creds: ", creds)
	service = build('calendar', 'v3', credentials=creds, cache_discovery=False)
	

	#response = service.events().get(calendarId='primary', eventId='eventId').execute()

	logging.info('Delete all the events previously created')

	if os.path.isfile(eventfilename):
		logging.info('File %s exists', eventfilename)
		with open(eventfilename, 'rb') as f:
			eventId = pickle.load(f)
		os.remove(eventfilename)
		logging.info('Removed file %s', eventfilename)
		for ids in eventId:
			service.events().delete(calendarId='primary', eventId=ids).execute()
		logging.info('Calendar events successfully deleted')
	else:
		print('Filename {} does not exist'.format(eventfilename))
	
		

	





## main file ----------------
def main():
	"""
	main file in action
	"""
	all_sem_dates = None
	site = None
	# clear screen operation
	if platform=='linux' or platform=='linux2' or platform=='darwin':


		os.system('clear')
	elif platform=='win32':
		os.system('cls')
	else:
		logging.info("Unknown OS.")
		exit(1)

	print("================= STARTING operation ====================")
	# main program starts 


	logging.info("[MAIN] Starting in the main.")


	# clearning option

	answer = input("Do you want to clean and fresh start? Yes: y/Y No:n/N?  ")
	

	if answer in ['y', 'Y']:
		logging.info('Cleanig all calendar data and files generated in previous runs')
		print('Starting fresh.\n\n')
		cleanUp(event_file_name)
	elif answer in ['n', 'N']:
		logging.info('Normal start')



	#print("Fall semester months: ", MONTHS_TO_CHECK_NAME)

	answer = input("Which schedule you want to follow? Press 1: UNO, 2: UNL, 3: UNMC, 4: Test     ")
	answer=int(answer)

	if answer == 1:
		site = uno_site
		tmp_file = uno_temp_file
		logging.info("[MAIN] UNO schedule chosen")
	elif answer == 2:
		site = unl_site
		tmp_file = unl_temp_file
		logging.info("[MAIN] UNL schedule chosen")
	elif answer == 3:
		site = unmc_site
		tmp_file = unmc_temp_file
		logging.info("[MAIN] UNMC schedule chosen")
	elif answer == 4:
		site = test_site
		tmp_file = test_temp_file
		logging.info("[MAIN] UNMC schedule chosen")
	else:
		logging.info("Not a valid option")

	answer = input("Please select which semester you want to follow: 1. Fall, 2. Spring:  ")
	answer=int(answer)
	
	if answer==1:
		
		# semester always start from Monday
		semester="FALL"
		
		SemDates = namedtuple('SemDates', 'year, month, day')



		sem_start = SemDates(year=year, month=fall_month, day=fall_day)
		sem_end = SemDates(year=year, month=fall_end_month, day=fall_end_day)
		#logging.info("Fall months are: %s", month for month in MONTHS_TO_CHECK_NAME)
		print("list all sem days: ")
		#printDateTime(listAllSemesterdays(sem_start, sem_end))
		all_sem_dates, tu_dict, th_dict = listAllSemesterdays(sem_start, sem_end)

		logging.info("[MAIN] working on extracting %s datetimes from string literals", semester)

	elif answer==2:
		pass
	else:
		print('Invalid info')

	holidays = checkHolidays(site, tmp_file)
	print("holidays: ", holidays)
	holidays = extractDateTimes(holidays, year, True)
	#print("Holidays this semester:", holidays)
	#print(date_header)
	#print("all_sem_dates", all_sem_dates)
	#print("Holidays this semester regex: ")
	#printDateTime(extractDateTimes(holidays, year))

	answer = input("Please select: 1. Excel, 2. calendar+excel  ")
	answer=int(answer)
	
	if answer == 1:
		logging.info("Just excel sheet is created .")
		createExcel(task_cols, all_sem_dates, xl_filename, holidays, False)
	elif answer == 2:
		logging.info("Both events in calendar and excel sheet are created .")
		#print(all_sem_dates)
		allEventsForCalendar = createExcel(task_cols, all_sem_dates, xl_filename, holidays, True)
		#print("allEventsForCalendar: ", allEventsForCalendar)
		createEventGoogleCalendar(allEventsForCalendar, event_file_name)

			

	else:
		print("Invalid option")

if __name__=="__main__":

	main()
