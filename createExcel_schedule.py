"""
	__author__="Subharthi"
	__date__=""
	__version__="v1.0"

	creates an excel file where it creates entries for
	class dates 

	Requirements: install beautifulsoup
"""
import logging
import xlsxwriter 
  

from sys import platform
import os
from datetime import date, timedelta 
from collections import namedtuple

import threading
import time
import urllib.request 
from bs4 import BeautifulSoup

import re

__author__="Subharthi"
__date__=""
__version__="v1.0"


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

labs_hw = ['Lab Assignment Posted', 'Lab Assignment Due', 
			'HW Assignment Posted', 'HW Assignment Due',
			'Special Notice']
date_header = "Class Dates"
isQuizToday = "Quiz Today"
isExamToday = "Exam Today"
task_cols.insert(0,date_header)
for data in labs_hw:
	task_cols.append(data)
task_cols_n = [i for i in range(0, len(task_cols))]
task_cols_dict = dict(zip(task_cols, task_cols_n))

# ====================== important dates ============================



# ===================================================================






if_quiz = False
if_meeting = False
if_holiday = False
if_exam=False
n_labs=5
n_hws = 3
lab_numbers = ["Lab Assignment - "+str(i) for i in range(1, n_labs+1)]
lab_description=["Implement Serial", "SPI Lab", "I2C lab",
				"Schematic Design", "PCB design"]

hw_description=["Initial Block Diagram", "Final Block Diagram", 
				"Component Selection"]

hw_numbers = ["HW Assignment - "+str(i) for i in range(1, n_labs+1)]

pcb_description = "PCB Submission for Manufacturing"
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

pcb_due = 8

project_demo_due=15
project_report_due=15

# ===================================================================
year=2019
fall_month=8
fall_day=26

fall_end_day=6
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

temp_file = "test_soup.txt"
uno_temp_file = "uno_test_soup.txt"
unl_temp_file = "unl_test_soup.txt"
unmc_temp_file = "unmc_test_soup.txt"


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
		logging.info("File does not exist, opening from internet %s", uno_site)
		url = urllib.request.urlopen(site)

		content = url.read()
		
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

	print(appendTextHolidays(holidays_with_reason, tmp_file))

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
		tuesday = sem_start_date + timedelta(days=i+1)
		thursday = sem_start_date + timedelta(days=i+3)
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
	

def createExcel(header, writeables, filename, holidays):
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
	#bold = workbook.add_format({'bold': True})
	#bg_blue = workbook.add_format({'bg_color': 'blue'})
	#bg_red = workbook.add_format({'bg_color': 'red'})
	#bg_yellow_bold = workbook.add_format({'bold':True,'bg_color': 'yellow'})
	for head in header:
		
		worksheet.write(row, col, head,  cell_format)
		col = col+1
	row = row + 1
	col = 0
	count = 2
	week = 0
	
	for ele in writeables:
		
		if count == 2:
			
			week = week + 1
			Week = "Week - " + str(week)
			hw_numbers_iter = iter(hw_numbers)
			lab_numbers_iter = iter(lab_numbers)
			format_cell = [True, 'white', 
				'black', '8', True]
			format_dict=createFormatExcel(format_cell)
			cell_format = workbook.add_format(format_dict)
			if week in hw_post_dates:
				text = hw_numbers[week-1]+'('+ hw_dict[hw_numbers[week-1]]+')'
				
				worksheet.write(row, task_cols_dict['HW Assignment Posted'],
								 text, cell_format)
			if week in lab_post_dates:
				text = lab_numbers[week-1]+'('+lab_dict[lab_numbers[week-1]] +')'
				worksheet.write(row, task_cols_dict['Lab Assignment Posted'],
								 text, cell_format)
			if week in lab_due_dates:
				text = lab_numbers[week-2]+'('+ lab_dict[lab_numbers[week-2]]+')'
								 
				worksheet.write(row, task_cols_dict['Lab Assignment Due'],
								 text
								 , cell_format)
			if week in hw_due_dates:
				text =  hw_numbers[week-2]+'('+hw_dict[hw_numbers[week-2]]+')'
				worksheet.write(row, task_cols_dict['HW Assignment Due'],
								text, cell_format)

			
			logging.info("Writing %s", Week)
			format_cell = [True, 'yellow', 
				'black', '10', True]
			format_dict=createFormatExcel(format_cell)
			cell_format = workbook.add_format(format_dict)
			worksheet.write(row, col, Week, cell_format)
			row = row + 1
			count = 0

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



def createAssignment(posted_dates, due_in, lab_assignment):
	"""
	rtype:
	"""
	pass

if __name__=="__main__":
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

	#print("Fall semester months: ", MONTHS_TO_CHECK_NAME)

	answer = input("Which schedule you want to follow? Press 1: UNO, 2: UNL, 3: UNMC     ")
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

	holidays = checkHolidays(site, tmp_file)
	
	holidays = extractDateTimes(holidays, year, True)
	#print("Holidays this semester:", holidays)
	#print(date_header)
	#print("all_sem_dates", all_sem_dates)
	#print("Holidays this semester regex: ")
	#printDateTime(extractDateTimes(holidays, year))

	
	createExcel(task_cols, all_sem_dates, xl_filename, holidays)
	