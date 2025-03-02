from datetime import datetime, timedelta
import json
import csv
import sys
from openpyxl import Workbook


def gettimeSpentAtWork(data):
    # load the list of dictionaries in the semanticSegments key
    semanticSegments = data['semanticSegments']
    print('semanticSegments size:', len(semanticSegments))

    # keep only the elements containing the key 'visit' with subkey 'topCandidate' with subkey 'semanticType' == 'INFERRED_WORK'
    csvData = []
    for segment in semanticSegments:
        if 'visit' in segment:
            visit = segment['visit']
            if 'topCandidate' in visit:
                topCandidate = visit['topCandidate']
                if topCandidate['semanticType'] == 'INFERRED_WORK':
                    start = segment['startTime']
                    end = segment['endTime']                
                    # subtract the start time from the end timestamp string
                    # by converting the string to a datetime object
                    start = datetime.strptime(start, '%Y-%m-%dT%H:%M:%S.%f%z')
                    end = datetime.strptime(end, '%Y-%m-%dT%H:%M:%S.%f%z')
                    
                    duration = end - start
                    # change date format for better readability
                    date = start.strftime('%Y-%m-%d')
                    start = start.strftime('%H:%M')
                    end = end.strftime('%H:%M')
                    
                    csvData.append([date, start, end, duration])
    
    csvData = combineEntries(csvData)

    return csvData

# function to combine double entries for the same day
def combineEntries(csvData):
    i = 0
    newList = []
    while i < len(csvData) - 1:
        if csvData[i][0] == csvData[i+1][0]:
            # combine the two entries
            date = csvData[i][0]
            start = csvData[i][1]
            end = csvData[i+1][2]
            # calculate the duration
            start_tmp = datetime.strptime(start, '%H:%M')
            end_tmp = datetime.strptime(end, '%H:%M')
            
            duration = end_tmp - start_tmp
            newList.append([date, start, end, duration])
            i += 2
        else:
            newList.append(csvData[i])
            i += 1
    return newList

# function to calculate the average duration each day
def calculateAverageDuration(csvData):
    acc = timedelta()
    for row in csvData:
        acc += row[3]
    average = acc / len(csvData)
    # convert to hours and minutes
    hours = int(average.total_seconds() // 3600)
    mins = int((average.total_seconds() % 3600) // 60)
    average_str = '%d:%02d' % (hours, mins)
    print('average duration in hours:Mins:', average_str)
    return average_str


def saveToCSV(csvData, filename):
    # save the data to a csv file
    with open(filename, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['Date', 'start', 'end', 'duration'])
        for row in csvData:
            # convert duration to string before writing to CSV
            row[3] = str(row[3])
            writer.writerow(row)
        return
    

def saveToExcel(csvData):    
    # save the data to a spreadsheet file
    wb = Workbook()
    ws = wb.active
    ws.title = 'Work'
    ws.append(['Date', 'start', 'end', 'duration'])
    for row in csvData:
        # convert duration to string before appending
        ws.append([row[0], row[1], row[2], str(row[3])])
    wb.save('work.xlsx')

def main():
    if len(sys.argv) < 2:
        print('Usage: python timetracker.py <input_file.json>')
        sys.exit(1)
    # Load json file
    with open(sys.argv[1]) as f:
        data = json.load(f)
    csvData = gettimeSpentAtWork(data)
    saveToCSV(csvData, 'work.csv')

if __name__ == '__main__':
    main()

