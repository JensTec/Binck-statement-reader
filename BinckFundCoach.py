#!/usr/local/bin/python3.9

# ------------ import functions ---------------------------------
import pdfplumber
from tabulate import tabulate
import os
import csv
from datetime import datetime
import shutil

# ------------ define global variables ---------------------------
rows = []
fields = ["TransID", "Transdate", "Rentedate", "Type", "Quantity", "Price", "Description", "Fund", "Position",
          "Debet", "Credit", "Statement date", "Filename", "Extra"]
path = "/Users/Sietse/Documents/Python projects/Binck_interpreter/pdfs"
files = []
skipped_files = []
skippath = str(path + "/Skipped")  # creates name and directory for skipped files
s = 0  # global skipped files counter


# ------------ directory / file 'reader'  ---------------------------
def directory_index():
    global path
    global files

    for x in os.listdir(path):  # show available csv and pdf files in path, and ability to select the correct one
        if x.endswith(".pdf"):
            files.append(x)


# ------------ extract data from table ---------------------------
def extracttable():
    global files
    global skipped_files
    global s
    empty = ['', '', '', '', '', '']
    regdate = '(\d{2}-\d{2}-\d{4})'  # regex string used by pdfplumber to search for the date in the statement
    table_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "text",
        "snap_y_tolerance": 5,
        "explicit_vertical_lines": {535},  # added due to the fact the last column was not automatically detected
    }
    k = 0  # counter to create unique transactions

    for fileID in files:
        with pdfplumber.open(path + "/" + fileID) as file:
            page0 = file.pages[0]  # list with data of the first page, to search the date
            searchdate = page0.search(regdate, regex=True, case=False)
            date = searchdate[0]['text']
            i = 0  # reset counter for number of pages

            try:
                for page in file.pages:
                    table = file.pages[i].extract_table(table_settings)  # extract tables from all possible pages
                    for row in table[3:]:
                        if row != empty:
                            if row[0] == '':
                                # part where description is managed
                                cleandescr = str(row[3]).replace("\n", " ")  # remove any \n from the descriptions
                                descr = rows[k-1][3] + " " + cleandescr
                                rows[k-1].pop(3)
                                rows[k-1].insert(3, descr)
                                # part where debit/credit amount is managed
                                if row[4] != '':
                                    rows[k-1].pop(4)
                                    rows[k-1].insert(4, float(".".join(row[4].replace(".", "").split(","))))
                                if row[5] != '':
                                    rows[k-1].pop(5)
                                    rows[k-1].insert(5, float(".".join(row[5].replace(".", "").split(","))))
                            else:
                                year = date[6:]  # extract year from statement date
                                transdate = row[1] + "-" + year  # create complete date including year
                                rentedate = row[2] + "-" + year  # create complete date including year
                                row.insert(1, transdate)  # insert complete date
                                row.pop(2)  # remove old date
                                row.insert(2, rentedate)  # insert complete date
                                row.pop(3)  # remove old date
                                row.insert(1, int(row[0]))  # add int transaction id
                                row.pop(0)  # remove string transaction id
                                row.insert(6, date)  # add statement date to row
                                row.insert(7, fileID)  # add FileID to row
                                rows.append(row)  # add row to list 'rows'
                                k += 1  # increase row counter
                    i += 1  # increase counter for number of pages
            except:
                skip = [date, fileID]
                skipped_files.append(skip)
                skipped_files.sort()
                if not os.path.exists(skippath):  # check if the skippath folder already exists
                    os.mkdir(skippath)
                shutil.move(path + "/" + fileID, skippath + "/" + fileID)  # move inputfile to 'skippath'
                s += 1

    # cut up description into usable parts
    for row in rows:
        descr_split = row[3].split()  # store complete description in a list, separated by spaces
        row.insert(3, descr_split[0])  # inserts type of transaction in position 3
        if descr_split[0] == "Koop" or descr_split[0] == "Verkoop":
            row.pop(4)  # remove original description

            # part to find and write the quantity of shares sold / purchased
            var1 = 1
            if descr_split[1].isdigit():  # search where the quantity is positioned
                var1 = 1
            elif descr_split[2].isdigit():
                var1 = 2
            elif descr_split[3].isdigit():
                var1 = 3
            elif descr_split[4].isdigit():
                var1 = 4
            elif descr_split[5].isdigit():
                var1 = 5
            if descr_split.count("bestens,") > 0:  # with 'bestens' is new way of formatting, without is old
                row.insert(4, descr_split[5])  # takes the quantity out of the description (position based)
            else:  # searches right position of the quantity of stocks purchased or sold
                row.insert(4, descr_split[var1])  # writes the quantity based on where it finds the digits

            # part to find and convert the price into a float
            tempprice1 = descr_split[(descr_split.index("@")+1)]  #
            price2 = tempprice1.rstrip(",")  # strip an extra , on the right side of the string if present
            price3 = price2.split(",")  # split price in two strings
            finalprice = float(".".join(price3))  # make it a float
            row.insert(5, finalprice)  # should show the price of the stock

            row.insert(6, " ".join(descr_split[1:]))  # insert the description
            row.insert(7, " ".join(descr_split[(var1 + 2):descr_split.index("@")]))  # insert fund name

            # part to find the position and write it as a float
            temp_pos = descr_split[descr_split.index("transactie:")+1].replace(".", "")
            if temp_pos.count(",") > 0:
                temp_pos2 = temp_pos.split(",")
                row.insert(8, float(".".join(temp_pos2)))  # position
            else:
                row.insert(8, float(temp_pos))  # position

        else:  # this is for "Verrekening", "Toekenning", "Uitkering"
            row.pop(4)
            row.insert(4, "")  # quantity
            row.insert(5, "")  # price
            row.insert(6, ' '.join(descr_split))  # description
            row.insert(7, "")  # fund
            row.insert(8, "")  # position

    rows.sort()  # sort rows with ascending transaction numbers
    print(tabulate(rows, headers=fields))  # printing the table on screen with the data
    print(tabulate(skipped_files))


# ------------ write csv file ---------- ---------------------------
def filecreation():
    global fields
    global path
    global rows
    a = datetime.now()  # unique timestamp
    filename = str(path + "/The Binck files " + str(a.year) + str(a.month).zfill(2) + str(a.day).zfill(2) + "_"
                   + str(a.hour).zfill(2) + str(a.minute).zfill(2) + ".csv")

    with open(filename, 'w') as csvfile:
        csvwriter = csv.writer(csvfile)  # creating a csv writer object
        csvwriter.writerow(fields)  # writing the fields
        csvwriter.writerows(rows)  # writing the data rows

    transactions = len(rows)

    print('\n')
    print("You have created the following file: '" + filename[len(path) + 1:] + "' with a total of %s transactions." %
          transactions)
    print(str(s) + " Files have been moved to the following folder: %s." % skippath)
    print('\n')


directory_index()


extracttable()


filecreation()


# how to debug pdfplumber
# img = p0.to_image()
# img.reset().debug_tablefinder(table_settings)
# img.show()
