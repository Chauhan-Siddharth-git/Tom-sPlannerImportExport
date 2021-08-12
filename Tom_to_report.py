import xlrd
import xlsxwriter
import csv
# import pandas as pd


NAME_OF_FILE_HERE = 'REPORT.xls'  # name of file where the report will be

input_file = "Tom_to_report_input.txt"


def input_file_read():

    stuff_in_file = ['THIS FILE MUST BE IN THE SAME DIRECTORY AS THE PROGRAM', '', 'DIRECTORY OF FIRST FILE', '',
                     'Pathway: ', '', 'PLEASE ENTER THE COLUMNS YOU WANT TO REMOVE', '',
                     'Remove: ', '', 'PLEASE ENTER THE WIDTHS OF THE REMAINING COLUMNS', '',
                     'Widths: ', '', 'DIRECTORY OF SECOND FILE', '', 'Pathway: ', '',
                     'PLEASE ENTER THE COLUMNS YOU WANT TO REMOVE', '', 'Remove: ', '',
                     'PLEASE ENTER THE WIDTHS OF THE REMAINING COLUMNS', '', 'Widths: ']

    try:
        with open(input_file) as f:
            reader = csv.reader(f, delimiter=' ')
            words = list(reader)

        return words

    except FileNotFoundError:
        f = open(input_file, "w")
        f.write('\n'.join(stuff_in_file))
        print("\nA file has been created in the same directory as the program", "\n\nName:\t", input_file)
        return []


def letter_to_num(x):
    stuff = x.upper()
    return ord(stuff) - ord('A')


def getting_an_array(file, TO_REMOVE1):
    loc = file  # the location of the file where we will read
    # the data from

    workbook = xlrd.open_workbook(loc)
    worksheet = workbook.sheet_by_index(0)

    for a in range(len(TO_REMOVE1)):
        TO_REMOVE1[a] = letter_to_num(TO_REMOVE1[a])

    a = []

    for i in range(worksheet.nrows):  # adds the data into the array 'a'
        a.append(worksheet.row_values(i))

    for count in range(len(a)):
        for u in range(len(a[count])):
            if u in TO_REMOVE1:  # Asks if the index is in the remove index
                a[count][u] = "TO POP"  # if it is, it is set equal to "TO POP"

    for y in range(len(a)):
        while "TO POP" in a[y]:  # Searches for "TO POP"
            a[y].remove("TO POP")  # then removes it

    for i in range(len(a)):
        for j in range(len(a[i])):
            if a[i][j] == "column description":  # Searches for "column description"
                a[i][j] = 'column'  # Then changes it
                # print("THERE HAS BEEN ONE INSTANCE")

    return a  # returns the array


def write(a, a1, WIDTHS1, WIDTHS2):

    workbook = xlsxwriter.Workbook(NAME_OF_FILE_HERE)
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Hello world')

    more = 0
    for i in range(len(a)):
        if a[i][0] == '2':  # checks if the beginning of the row starts with a 2 so it can add +1 to more
            # print(a[i][0])          # in order to leave a gap between rows that start with 1 and 2
            more = more + 0
            worksheet.write_row(i, 0, '')
        for j in range(len(a[i])):
            worksheet.write(i + more, j, a[i][j])

    american = workbook.add_format(
        {'num_format': 'mm/dd/yy'})  # New format called american because its Month-Day-Year lol

    start = 0  # placeholder for the start date column
    last = 0  # placeholder for the end date column
    for i in range(
            len(a[0])):  # finds the start and end date columns and stores them in the placeholders 'start' and 'last'
        if a[0][i] == "start row":
            start = i

        if a[0][i] == "end row":
            last = i
    # print("start:", start, "end:", last)

    for date in range(2, len(a)):  # Adds the formatting on the start and end dates columns
        worksheet.write(date, start, a[date][start],
                        american)  # Overwrites whatever is in the cell with the same value with a new format.
        worksheet.write(date, last, a[date][last], american)  # Same here
        # print("start:", date, start, "end:", date, last)

    values = chart(a)
    distance = 2

    color1 = workbook.add_format({'bg_color': '#AAFFAA'})

    bold = workbook.add_format({'bold': 1})  # Bold formatting for the chart title

    # Add the worksheet data that the charts will refer to.
    headings = ['Category', 'Values']
    data = values

    # worksheet.write_row('R3', headings, bold)
    worksheet.write(2, len(a[0]) + distance, headings[0], bold)
    worksheet.write(2, len(a[0]) + distance + 1, headings[1], bold)
    # worksheet.set_row('R1', cell_format=color1)
    worksheet.write_column(3, len(a[0]) + distance, data[0])
    worksheet.write_column(3, len(a[0]) + distance + 1, data[1])

    #######################################################################
    #
    # Create a new chart object.
    #
    chart1 = workbook.add_chart({'type': 'doughnut'})

    # Configure the series. Note the use of the list syntax to define ranges:
    chart1.add_series({
        'name': 'Pie sales data',
        'categories': ['Sheet1', 3, len(a[0]) + distance, 5, len(a[0]) + distance],
        'values': ['Sheet1', 3, len(a[0]) + distance + 1, 5, len(a[0]) + distance + 1],
        'points': [
            {'fill': {'color': '#376092'}},
            {'fill': {'color': '#95b3d7'}},
            {'fill': {'color': '#b9cde5'}},
        ],
    })

    # Add a title.
    chart1.set_title({'name': 'Portfolio Execution'})

    chart1.set_rotation(330)

    # Insert the chart into the worksheet (with an offset).
    worksheet.insert_chart(2, len(a[0]) + distance + 2, chart1, {'x_offset': 25, 'y_offset': 0})

    # Array of columns along with their width size

    for i in range(int(len(WIDTHS1) / 2)):  # goes through the 'widths' array
        # x = letter_to_num(WIDTHS1[i * 2])  # Converts the columns to a number corresponding to the column
        x = i
        # print("THIS IS THE WIDTH: ", x, int(WIDTHS1[(i * 2)+1]))
        if int(WIDTHS1[(i * 2) + 1]) == 0:
            worksheet.set_column(x, x, None, None, {'hidden': 1})
            # print(int(WIDTHS1[(i * 2) + 1]))
        else:
            worksheet.set_column(x, x, int(WIDTHS1[(i * 2) + 1]))  # sets the column width

    worksheet.set_row(1, cell_format=bold)
    worksheet.set_row(1, cell_format=color1)

    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Hello world')

    for i in range(len(a1)):
        for j in range(len(a1[i])):
            worksheet.write(i + more, j, a1[i][j])

    for i in range(int(len(WIDTHS2) / 2)):  # goes through the 'widths' array
        # x = letter_to_num(WIDTHS1[i * 2])  # Converts the columns to a number corresponding to the column
        x = i
        # print("THIS IS THE WIDTH: ", x, int(WIDTHS2[(i * 2)+1]))
        if int(WIDTHS2[(i * 2) + 1]) == 0:
            worksheet.set_column(x, x, None, None, {'hidden': 1})
            # print(int(WIDTHS2[(i * 2) + 1]))
        else:
            worksheet.set_column(x, x, int(WIDTHS2[(i * 2) + 1]))  # sets the column width

    worksheet.set_row(1, cell_format=bold)
    worksheet.set_row(1, cell_format=color1)

    workbook.close()  # don't wanna be irresponsible now do we?


def chart(a):
    stats = "Status"  # The column we wanna look at
    begin = 0  # the placeholder for "Status" column
    # distance = 3

    for i in range(len(a[0])):
        if a[0][i] == stats:
            begin = i  # sets placeholder when it finds the status column
    # print(begin)
    values = [0, 0, 0]  # placeholder for the data
    to_find = [["Completed"], ["In progress"], ["Not started"]]  # things to find within the "status" column

    for stuff in range(2, len(a)):

        # checks to see if it is only performing tasks on rows starting with 1, if not then
        # it breaks out of the loop

        if a[stuff][begin] in to_find[0]:
            values[0] += 1

        if a[stuff][begin] in to_find[1]:
            values[1] += 1

        if a[stuff][begin] in to_find[2]:
            values[2] += 1

    # print(values)

    to_send = [[], []]

    to_send[0] = ["Completed", "In progress", "Not Started"]  # The labels for the data
    to_send[1] = [values[0], values[1], values[2]]  # The data
    return to_send  # returns the labels and the data


def instructions():
    print("\nA FILE HAS BEEN CREATED/UPDATED CALLED", NAME_OF_FILE_HERE)


def main():

    if len(input_file_read()) > 1:

        load = input_file_read()
        excel_file_one = load[4][1]
        excel_file_two = load[16][1]
        TO_REMOVE1 = load[8][1:len(load[8])]
        TO_REMOVE2 = load[20][1:len(load[20])]
        WIDTHS1 = load[12][1:len(load[12])]
        WIDTHS2 = load[24][1:len(load[24])]

        # print("THIS IS THE TEST DATA: ", excel_file_one)
        array1 = getting_an_array(excel_file_one, TO_REMOVE1)
        array2 = getting_an_array(excel_file_two, TO_REMOVE2)
        # print("THIS IS ARRAY 2", array2)
        write(array1, array2, WIDTHS1, WIDTHS2)
        instructions()


if __name__ == '__main__':
    main()
