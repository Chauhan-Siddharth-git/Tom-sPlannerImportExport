import time
import pandas as pd
import os.path
from os import path
from pathlib import Path
import sys
import xlrd
import csv
import datetime
import re

input_file_name = "report_to_Tom_input.txt"

column = ''
out_file = "OUTPUT.txt"
names = ['PROJECT PORTFOLIO', 'DEMANDS']


def letter_to_num(x):
    stuff = x.upper()
    return ord(stuff) - ord('A')


def file(daName):
    name = daName
    # folder = Path(dir)
    # file_to_open = folder / name
    file_to_open = name
    send = []
    stuff_in_file = ['PLEASE ENTER THE PATHWAY OF THE EXCEL FILE YOU WANT TO READ', '', 'Pathway: ', '',
                     'PLEASE ENTER THE COLUMN YOU WANT TO READ UP TO FOR THE FIRST SHEET (NON INCLUSIVE)', '',
                     'Column: ', '',
                     'PLEASE ENTER THE COLUMN YOU WANT TO READ UP TO FOR THE SECOND SHEET (NON INCLUSIVE)', '',
                     'Column: ', '', ]  # stuff that should be in the file

    try:  # this works if the file exists
        with open(file_to_open) as f:
            reader = csv.reader(f, delimiter=' ')
            words = list(reader)
            # print(words)
            send = words

        # print("\nInput file has been opened")
        return send

    except FileNotFoundError:  # this works if the file doesn't exist
        f = open(file_to_open, "w")
        f.write('\n'.join(stuff_in_file))
        print("\nA file has been created in the same directory as the program", "\n\nName:\t", name)
        return send


def format(stuff):  # this function converts an array with rows * columns into columns * rows, like a 90 degree turn
    output = []

    for i in range(len(stuff[0])):
        temp = []
        for j in range(len(stuff)):
            temp.append(stuff[j][i])

        output.append(temp)
    # print(output)
    return output


def write(stuff, num, cutoff):  # this copies data onto your clipboard

    cutoff = letter_to_num(cutoff)
    # print("THIS IS THE CUTOFF:", cutoff)
    write = []

    for i in range(len(stuff)):
        temp = []
        for j in range(cutoff):
            temp.append(stuff[i][j])
        write.append(temp)

    # for i in range(len(write)):
    #    print(i, write[i])

    df = pd.DataFrame(write)
    # print(df)
    new_header = df.iloc[0]

    df = df[1:]

    df.columns = new_header
    # df.to_clipboard(sep='	')
    # df.to_clipboard(sep='	', index=False)

    file = open(out_file, 'a')

    # print("\n\n", name_with_date(num))
    file.write("\n\n" + name_with_date(num))
    file.write("\n\n")

    df.to_csv(out_file, sep='	', index=False, mode='a')

    file.write("\n\nNEW SHEET")
    for i in range(3):
        file.write("\n//////////////////////////////////////////////////////////////////////////////////////////////////////")
    file.write("\nNEW SHEET\n\n")
    file.close()


def convert_date(stuff):  # converts from YYYY-MM-DD to MM-DD-YYYY
    output = stuff[5:10] + stuff[4:5] + stuff[0:4]
    return output


def mmddyy(stuff, what_kind):
    convert = stuff
    for i in range(len(convert)):
        for j in range(len(convert[i])):
            if type(convert[i][j]) is what_kind:  # looks for type Timestamp and converts the date inside
                # print('yes')
                convert[i][j] = convert_date(str(convert[i][j]).split(' ')[0])  # converts it to mm-dd-yyyy

    return convert


def name_with_date(num):
    e = datetime.datetime.now()
    # print("Today's date:  = %s/%s/%s %s:%s:%s" % (e.month, e.day, e.year, e.hour, e.minute, e.second))
    name = ("%s/%s/%s-%s:%s:%s" % (e.month, e.day, e.year, e.hour, e.minute, e.second))
    # print("PORTFOLIO", name)
    return names[num] + ", " + name


def main():
    # stuff = read()

    f = open(out_file, 'w+')  # Erases the output file to update
    f.seek(0)
    f.truncate()
    f.close()

    if len(file(input_file_name)) > 1:
        info = file(input_file_name)
        columns = []
        read_file = info[2][1]

        columns.append(info[6][1])
        columns.append(info[10][1])

        xl = pd.ExcelFile(read_file)
        res = len(xl.sheet_names)

        for i in range(res):
            data = pd.read_excel(read_file, sheet_name=i)  # reads the data

            lis = data.values.tolist()

            try:
                # print(type(lis[5][11]))
                lis = mmddyy(lis, 'pandas._libs.tslibs.timestamps.Timestamp')

            except:
                pass

            write(lis, i, columns[i])
        print("\nOUTPUT FILE HAS BEEN CREATED/UPDATED CALLED:", out_file)


if __name__ == "__main__":
    main()
