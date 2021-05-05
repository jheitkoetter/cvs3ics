#!/usr/bin/env python3
#
# cvs3ics.py - Copyright (c) 2021 Jörg Heitkötter (joke)
#
# Sys is all we need: csv, ics, dateparser, argparser (and love)
import sys
import csv
from ics import Calendar, Event
import dateparser
import argparse

# parse args and show usage
parser = argparse.ArgumentParser()
parser.add_argument('input', type=str, help='input CSV file containing calendar events')
parser.add_argument('output', type=str, help='output iCal/ICS file')
args = parser.parse_args()

# WARN
def warning(*objs):
    print("WARNING: ", *objs, file=sys.stderr)

# FAIL
def fail(*objs):
    print("FATAL: ", *objs, file=sys.stderr)
    sys.exit(1)
    
# the working horse
def csv2ical(input_file, output_file):
  # Convert a CSV file with event information to ical,
  # for import to Google Calendar, Microsoft Outlook and etc.
  # 
  # CSV format https://support.google.com/calendar/answer/37648
  # Subject,Start Date,Start Time,End Date,End Time,All Day Event,Description,Location,Private
  # 
  # usage: cvs3ics.py <CSV-input-file> <ICS-output-file>
  #
  
  # open and read CSV file
  with open(input_file) as csv_file:
    reader = csv.DictReader(csv_file)
    #reader = csv.reader(csv_file)

    # create new calendar
    c = Calendar()

    # test if all columns needed are present
    for field in reader.fieldnames:
      if not field in ['Subject','Start Date','End Date','Start Time','End Time','All Day Event','Description','Location','Private']:
          warning("column '%s' is ignored" % field)

    for field in ['Subject','Start Date','End Date','Start Time','End Time','All Day Event','Description','Location','Private']:
      if not field in reader.fieldnames:
          fail("missing column '%s'" % field)

    # iterate over each line
    for n, row in enumerate(reader):
      # Google calendar has 2-9 columns
      if len(row) < 1:
        fail("ICS file should have 2 to 9 columns, ie. Subject,Start Date[,Start Time,End Date,End Time,All Day Event,Description,Location,Private]")
        
      # skip header
      if n == 0:
        continue

      summary = row['Subject']
      
      # order in dates may be ambiguous, and neither locale nor language help, so define what the order is:
      dtstart = dateparser.parse(row['Start Date']+' '+row['Start Time'], settings={'DATE_ORDER': 'DMY'})
      dtend = dateparser.parse(row['End Date']+' '+row['End Time'], settings={'DATE_ORDER': 'DMY'})
      
      # skip allday flag
      allday = row['All Day Event'].strip()
      description = row['Description'].strip()
      location = row['Location'].strip()
      # skip private flag
      private = row['Private'].strip()
      
      # create event
      e = Event(
            name=summary,
            begin=dtstart,
            end=dtend,
            description=description,
            location=location,
        )
      c.events.add(e)
      # print (e)

    # write to output file
    with open(output_file, 'w') as out_f:
      out_f.writelines(c)
      out_f.close()

# parse args & execute
def main(args):
  csv2ical(args.input, args.output)
  
if __name__ == "__main__":
  main(parser.parse_args())