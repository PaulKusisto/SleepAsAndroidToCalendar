import os
import csv
import argparse
from datetime import datetime

def main(inpath,outpath):
    iCalOpenLines = ["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//Paul Kusisto//SleepAsAndroidToCalendar.py 1.0.0//EN","X-WR-CALNAME:Sleep","X-WR-CALDESC:Imported sleep logs from Sleep As Android.\n"]
    iCalCloseLines = ["END:VCALENDAR"]
    
    sleepTuples = []
    
    # ingest data
    with open(inpath,'r') as sleepData:
        dataReader = csv.reader(sleepData,dialect='excel') # excel dialect seems to work well
        for row in dataReader:
            if row[0].isdigit():
                # this row has a proper, numeric id
                startDateTime = stringToDateTime(row[2])
                endDateTime = stringToDateTime(row[3])
                comment = row[7]
                sleepTuples.append((startDateTime, endDateTime, comment))
    
    # output iCal file
    with open(outpath,'w') as outFile:
        def writeSleepEvent(sleepTuple):
            # fuction to write an event to the open .ics file
            outFile.write("BEGIN:VEVENT\n")
            outFile.write("DTSTART:%s\n" % dateTimeToCalString(sleepTuple[0]))
            outFile.write("DTEND:%s\n" % dateTimeToCalString(sleepTuple[1]))
            outFile.write("DESCRIPTION:%s\n" % sleepTuple[2])
            #outFile.write("LOCATION:%s")
            outFile.write("SUMMARY:Sleep\n")
            outFile.write("END:VEVENT\n")
        
        # write the header
        outFile.write('\n'.join(iCalOpenLines))
        for sleepTuple in sleepTuples:
            writeSleepEvent(sleepTuple)
        outFile.write('\n'.join(iCalCloseLines))

def stringToDateTime(stringToParse):
    # Returns the datetime object represented by a string of the common format of Sleep As Android data
    return datetime.strptime(stringToParse, '%d. %m. %Y %H:%M')

def dateTimeToCalString(dateTimeToParse):
    # Returns a string representing the given DateTime object in iCal format
    return dateTimeToParse.strftime('%Y%m%dT%H%M%S%Z')

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description = "Simple tool to create a iCal file of all sleep records in a Sleep As Android backup file.")
    parser.add_argument('-i', '--input', dest="inputPath", default=os.path.dirname(os.path.realpath(__file__))+os.sep+"Sleep as Android Data.csv", help="Path of the Sleep As Android cloud backup file.  By default, the program will look for a file named 'Sleep as Android Data.csv' in the current directory.")
    parser.add_argument('-o', '--output', dest="outputFilePath", default=os.path.dirname(os.path.realpath(__file__))+os.sep+"sleepCal.ics", help="Directory where the final output will be saved.  By default, output will be saved to a file named 'sleepCal.ics' in the current directory.")
    args = parser.parse_args()
    main(args.inputPath,args.outputFilePath)
