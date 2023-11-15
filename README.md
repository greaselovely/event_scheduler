# Event Scheduler

Created this for the purpose of generating a social calendar for work and wrote it to be used by anyone to throw together a social calendar.

I created this for a breakfast and a happy hour event schedule.  This was largely an academic application, and to avoid my brain catching on fire from having to figure out what dates to schedule and what location to schedule.  

You'll have to create a list of destinations to go to.  Run address_generator.py and then you can update the fake demo data in the json files.

Once you have the json files created, you can edit event_scheduler.py to change the start and end dates and times.  Run it.  This will create a CSV, a XLSX, and an ics file for importing.

### Install Needed Modules

`pip install -r requirements.txt`
