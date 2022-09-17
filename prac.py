 # Reading an excel file using Python
import xlrd

# Give the location of the file
loc = ("MEC_Competition_Voting_Data.xls")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
rows = sheet.nrows
voters = []
validRows = []
blacklist = []

def getVotes():
    for j in range(rows):

        if (sheet.cell_value(j, 0) + sheet.cell_value(j, 1)) in voters:

            # removeIndex = voters.index((sheet.cell_value(j, 0) + sheet.cell_value(j, 1)))
            voters.remove((sheet.cell_value(j, 0) + sheet.cell_value(j, 1)))
            blacklist.append(sheet.cell_value(j, 0) + sheet.cell_value(j, 1))
            continue

        elif (sheet.cell_value(j, 0) + sheet.cell_value(j, 1)) in blacklist:
            continue

        else:
            if ',' not in sheet.cell_value(j, 2):
                voters.append(sheet.cell_value(j, 0) + sheet.cell_value(j, 1))
                validRows.append(j)
    PronouncedJiffUnion = 0
    PineappleOnPizza = 0
    SocksAndCrocsReformLeague = 0
    for k in validRows:
        if (sheet.cell_value(k, 2) == "Pineapple Pizza Party"):
            PineappleOnPizza = PineappleOnPizza + 1
        if (sheet.cell_value(k, 2) == "Pronounced Jiff Union"):
            PronouncedJiffUnion = PronouncedJiffUnion + 1
        if (sheet.cell_value(k, 2) == "Socks and Crocs Reform League"):
            SocksAndCrocsReformLeague = SocksAndCrocsReformLeague + 1

    return [PineappleOnPizza,PronouncedJiffUnion,SocksAndCrocsReformLeague]

 # checking for repeats of names

getVotes()












