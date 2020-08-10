import csv

data = []

with open ('fictData.csv', 'r') as csvfile:
    csvreader = csv.DictReader(csvfile)

    with open ('getRgs_output.txt', 'w') as txtFile:
        counter = 1
        for row in csvreader:
            tmp = {}
            tmp['index'] = counter
            counter += 1
            tmp['ra'] = row['RA']
            data.append(tmp)
        txtFile.write(str(data))