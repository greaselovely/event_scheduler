from icalendar import Calendar, Event
from datetime import datetime, timedelta
import pathlib

"""
Takes a list of lists, iterates and creates ical files.
Using this as an example.  This is strictly utility for 
previous needs, don't use it unless you need the ical files again.
"""

ical_path = pathlib.Path(__file__).parent / "ical"
pathlib.Path.mkdir(ical_path, exist_ok=True)

events = [
	["07/11/2023","Tuesday","07:00 AM","North General Consulting","1838 Sam Rds, Montgomery, AL 36105"],
	["08/01/2023","Tuesday","07:00 AM","People Power Hill","4994 Christal Xrd, Montgomery, AL 36117"],
	["08/22/2023","Tuesday","07:00 AM","West People Construction","9319 David Grns, Medford, MA 02155"],
	["09/12/2023","Tuesday","07:00 AM","Interactive Provider Advanced","6769 Alice Cmns, Richmond, VT 05477"],
	["10/03/2023","Tuesday","07:00 AM","Signal Studio Contract","8661 Annemarie Rds, San Leandro, CA 94578"],
	["10/24/2023","Tuesday","07:00 AM","Interactive Provider Advanced","6769 Alice Cmns, Richmond, VT 05477"],
	["11/14/2023","Tuesday","07:00 AM","Signal Studio Contract","8661 Annemarie Rds, San Leandro, CA 94578"],
	["12/05/2023","Tuesday","07:00 AM","Studio Contract Consulting","8711 Ahmed Wall, Anchorage, AK 99502"],
	["12/26/2023","Tuesday","07:00 AM","North General Consulting","1838 Sam Rds, Montgomery, AL 36105"],
	["01/16/2024","Tuesday","07:00 AM","Studio Contract Consulting","8711 Ahmed Wall, Anchorage, AK 99502"]
	]

def create_ics_file(eventlist):
	location = eventlist[3]
	address = eventlist[-1]
	date_pattern = "%m/%d/%Y %I:%M %p"
	event_date = eventlist[0] + " " + eventlist[2]
	event_start = datetime.strptime(event_date, date_pattern)
	event_end = event_start + timedelta(hours=1)
	ical_filename = f'{location}_{event_start.strftime("%m-%d-%Y")}'
	ical_file_path = pathlib.Path.joinpath(ical_path, f'{ical_filename}.ics')

	cal = Calendar()
	event = Event()
	event.add('summary', location)
	event.add('location', address)
	event.add('dtstart', event_start)
	event.add('dtend', event_end)
	cal.add_component(event)
	with open(ical_file_path, 'wb') as file:
		file.write(cal.to_ical())
	return ical_file_path

for event in events:
	print(create_ics_file(event))