import csv
import os
from operator import itemgetter
import pyexcel

# all text is NOT confidential, just thought I should mention that

# hi strager! this is a very messy piece of code that aims to take the data in 'ReportUnfixed.csv' and turn it into 'ReportIdealOutput.xlsx'
# currently I have managed to get it as far as 'ReportFixed.xlsx'
# I could use your help with two things!:
# 1. can you let me know how I can possibly make the below code cleaner? at the moment I bet it is a pain to look at
# 2. how could I reach the ideal output that is shown in 'ReportIdealOutput.txt'?
# thank you as always!

# remove lines that do not include 'Alpha' and lines that do include 'No Show'
with open('ReportUnfixed.csv', 'r') as imp:
    lines = imp.readlines()

with open('Report.csv', 'w', newline='') as out:
    for line in lines:
        if 'Alpha' in line and not 'No Show' in line:
            out.write(line)

# remove unnecessary columns
with open('Report.csv', 'r') as imp:
    reader = csv.reader(imp)
      
    with open('Report2.csv', 'w', newline='') as out:
        writer = csv.writer(out)
        for r in reader:
            writer.writerow((r[1], r[2], r[3]))

# sort name column to aplhabetical order
with open('Report2.csv', 'r') as imp:
    data = [line for line in csv.reader(imp)]

data.sort(key=itemgetter(1))

with open('ReportFixed.csv', 'w', newline='') as out:
    csv.writer(out).writerows(data)

# rename csv to xlsx for further formatting in a better module?
pyexcel.save_as(file_name='ReportFixed.csv', dest_file_name='ReportFixed.xlsx')

# cleanup
os.remove('Report.csv')
os.remove('Report2.csv')
os.remove('ReportFixed.csv')