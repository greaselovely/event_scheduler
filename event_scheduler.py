from datetime import datetime, timedelta
import random
import pandas as pd
import tabulate
import pathlib
from icalendar import Calendar, Event
from openpyxl import load_workbook
import json
import os

####
####
bf_start_date_and_time = "11-01-2023 0700"
hh_start_date_and_time = "11-08-2023 1600"

bf_repeat = 3	# number of weeks to repeat breakfast
hh_repeat = 2	# number of weeks to repeat happy hour
num_weeks = 10	# maximum number of weeks to display

breakfast_filename = "breakfast.json"	# input dict files
happyhour_filename = "happyhour.json"	# event dict files
events_csv = "events.csv"	# tab delimited file name
events_excel = "events.xlsx"	# excel file name
sheet_name = "Event Schedule"	# excel sheet name
ics_calendar_file = "BFHH_Schedule"	# combined ics file for calendar import, leave the extension off, we do it below
####
####


"""
This script reads in text files that contain a json 
of "regions" (north, south) with location names and associated
addresses to create an event schedule based on the repeat of 
each type of event, and randomizes the location.

If you need an example of a dictionary to write, run address_generator.py
and you'll see how the files are structured.  

Modify this to use these dates and times to create a schedule
bf = breakfast hh = happy hour
"""




local_path = pathlib.Path(__file__).parent / "event_files"
local_path.mkdir(parents=True, exist_ok=True)
breakfast_fullpath = pathlib.Path.joinpath(local_path, breakfast_filename)
happyhour_fullpath = pathlib.Path.joinpath(local_path, happyhour_filename)
excel_fullpath = pathlib.Path.joinpath(local_path, events_excel)

if not pathlib.Path.is_file(breakfast_fullpath) and not pathlib.Path.is_file(happyhour_fullpath):
	print("\nOne or both event files do not exist, exiting...\n\n")
	exit()

with open(breakfast_fullpath, 'r') as bfile:
	happy_hour = json.loads(bfile.read())

with open(happyhour_fullpath, 'r') as hfile:
	breakfast = json.loads(hfile.read())

def clear():
    os.system("cls" if os.name == "nt" else "clear")


def get_dates_every_n_weeks(start_date: str, repeat: int, num_weeks: int) -> None:
	"""
	This takes in a date formatted string, and two integers and generates
	a list of dates for every 'repeat' number of weeks and creates a list
	"""
	dt_pattern = "%m-%d-%Y %H%M"
	current_date = datetime.strptime(start_date, dt_pattern)
	dates.append(current_date)
	for _ in range(num_weeks - 1):
		current_date += timedelta(weeks=repeat)
		dates.append(current_date)


def generate_calendar(d: dict) -> pd.DataFrame:
	"""
	This is passed a dictionary to reference and randomly chooses locations,
	creates a list for each column, creates a dataframe from the lists, 
	passed to tabulate for print to stdout,
	and calls a second function for each event to generate the icalendar files
	We also create a header and to each dataframe iteration to make viewing easier.
	
	Returns a dataframe.  
	"""
	events_day = []
	events_date = []
	events_time = []
	events_location = []
	events_address = []
	for i, date in enumerate(dates):
		if i % 2 == 0:
			location, address = random.choice(list(d.get("north").items()))
		else:
			location, address = random.choice(list(d.get("south").items()))
		events_day.append(date.strftime("%A"))
		events_date.append(date.strftime("%m/%d/%Y"))
		events_time.append(date.strftime("%I:%M %p"))
		events_location.append(location)
		events_address.append(address)

		create_ics_file(date, events_location[i], events_address[i])
	
	df = pd.DataFrame({"Date" : events_date, "Day" : events_day, "Time": events_time, "Location" : events_location, "Address" : events_address})
	print(tabulate.tabulate(df, showindex=False, headers=df.columns))

	blank_list = [None]	# need None in there to actually make the dataframe complete with blank "stuff"
	blank_df = pd.DataFrame({"Date" : blank_list, "Day" : blank_list, "Time": blank_list, "Location" : blank_list, "Address" : blank_list})
	if d == breakfast: 
		header_df = pd.DataFrame({"Date" : "Breakfast", "Day" : blank_list, "Time": blank_list, "Location" : blank_list, "Address" : blank_list})
	else:
		header_df = pd.DataFrame({"Date" : "Happy Hour", "Day" : blank_list, "Time": blank_list, "Location" : blank_list, "Address" : blank_list})
	
	concat_df = [header_df, df, blank_df]
	final_df = pd.concat(concat_df, axis=0, join='outer', ignore_index=True)
	return final_df


def create_csv_and_excel(dataframe: pd.DataFrame) -> tuple:
	"""
	Takes a dataframe and saves it to a csv, which is tab delimited 
	from the dataframe and also creates an xlsx.  Calles another function
	adjust column widths to clean up the spreadsheet for viewing
	"""
	csv_path = pathlib.Path.joinpath(local_path, events_csv)
	excel_fullpath = pathlib.Path.joinpath(local_path, events_excel)
	dataframe.to_csv(csv_path, index=False, sep='\t', mode='w')
	dataframe.to_excel(excel_fullpath, index=False, sheet_name=sheet_name)

	adjust_column_widths(excel_fullpath)
	
	return csv_path, excel_fullpath


def adjust_column_widths(filename: str) -> None:
    """
	Used to adjust the width of each column to make it easier to read
	called from the function above.
	"""
    workbook = load_workbook(filename)
    sheet = workbook.active
    column_widths = {'A': 15, 'B': 15, 'C': 15, 'D': 25, 'E': 52 }
    for column, width in column_widths.items():
        sheet.column_dimensions[column].width = width
    workbook.save(filename)


def create_ics_file(date: datetime, location: str, address: str):
	"""
	Takes a datetime object, location and address as strings
	and creates an individual icalendar file saved as 
	location_date.ics : ie... Restaurant_01-01-2001.ics
	"""
	global ical_path
	ical_path = local_path / 'ical'
	ical_path.mkdir(parents=True, exist_ok=True)

	cal = Calendar()
	event = Event()
	event.add('summary', location)
	event.add('location', address)
	events_start = date
	event.add('dtstart', events_start)
	events_end = events_start + timedelta(hours=1)
	event.add('dtend', events_end)
	cal.add_component(event)
	ical_filename = f'{location}_{date.strftime("%m-%d-%Y")}'
	ical_file_path = pathlib.Path.joinpath(ical_path, f'{ical_filename}.ics')
	with open(ical_file_path, 'wb') as file:
		file.write(cal.to_ical())
	
def generate_combined_ics():
	"""
	The original version did not contain this function as I hadn't
	thought about making a single file to upload the calendar.
	This will iterate over the ical directory, combine all of the file
	contents into a single file and then delete each individual file 
	after processing.
	"""
	today = datetime.now()
	output_filename = f'{ics_calendar_file}_{today.strftime("%m-%d-%Y")}.ics'
	ical_path_obj = pathlib.Path(ical_path)
	output_file_path = local_path.joinpath(output_filename)
	combined_calendar = Calendar()

	ics_files = [f for f in ical_path_obj.glob("*.ics")]

	for ics_file in ics_files:
		with open(ics_file, "rb") as file:
			calendar_data = file.read()
			calendar = Calendar.from_ical(calendar_data)

			for component in calendar.walk("VEVENT"):
				combined_calendar.add_component(component)

		os.remove(ics_file)

	with open(output_file_path, "wb") as output_file:
		output_file.write(combined_calendar.to_ical())
	
	ical_path.rmdir()



def main():
	clear()
	global dates

	what_to_schedule = input("Breakfast and Happy Hour? (y/n) [y]: ")
	what_to_schedule = what_to_schedule.lower()
	if what_to_schedule == "": what_to_schedule = "y"

	if what_to_schedule == "y":
		dates = []
		print("\n\n  Breakfast Schedule")
		get_dates_every_n_weeks(bf_start_date_and_time, bf_repeat, num_weeks)
		bf_df = generate_calendar(breakfast)

	dates = []	# clears previous list for happy hour
	print("\n\n  Happy Hour Schedule")
	get_dates_every_n_weeks(hh_start_date_and_time, hh_repeat, num_weeks)
	hh_df = generate_calendar(happy_hour)

	"""
	We are creating a list of both dataframes above, and doing a concat on
	the two dataframes unless we are only doing happy_hour, ignoring the indexes
	and then shipping them off for file creation (csv, xlsx)
	"""
	
	if what_to_schedule == "y":
		concat_df = [bf_df, hh_df]
	else:
		concat_df = [hh_df]

	final_df = pd.concat(concat_df, axis=0, join='outer', ignore_index=True)

	c,x = create_csv_and_excel(final_df)
	print(f"\n\nFiles Saved :\n{c}\n{x}\n\n")
	generate_combined_ics()


if __name__ == "__main__":
	main()



