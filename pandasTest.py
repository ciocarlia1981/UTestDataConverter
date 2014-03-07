__author__ = '38473'

import pandas as pd
import os
import xlrd

allParticipants = pd.DataFrame()

for file in os.listdir("data sheets"):
    try:
        print "0---"
        print file

        book = xlrd.open_workbook("data sheets/"+file)

        backgroundSheet = book.sheet_by_name("Background")
        headerRow = 0

        for rowNumber in range(2, backgroundSheet.nrows):
            if backgroundSheet.cell_value(rowNumber, 0) == "P#":
                headerRow = rowNumber

        print backgroundSheet.row(headerRow)

    except Exception as error:
        print error

#allParticipants.to_csv("test.csv", sep="\t", encoding="utf-8")