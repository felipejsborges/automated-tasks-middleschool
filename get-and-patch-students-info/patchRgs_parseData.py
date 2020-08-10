import csv

data = []

with open ('fictData.csv', 'r') as csvfile:
    csvreader = csv.DictReader(csvfile)

    with open ('patchRgs_output.txt', 'w') as txtFile:
        for row in csvreader:
            tmp = {}
            tmp['ra'] = row['RA']
            tmp['rg'] = row['RG']
            tmp['dig'] = row['DIG']
            tmp['uf'] = 'SP'
            tmp['emis'] = row['EXP']
            data.append(tmp)
        txtFile.write(str(data))