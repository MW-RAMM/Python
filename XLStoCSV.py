import xlrd  # required for excel functions
import sys  # required for passing in argyments

# lExcel_FileName = "D:\\Development\\Python\\XLStoCSV\\ExcelSource\\DailyTransactions2019829782539a7-8520-4185-a36a-9b3e415ae22f.xls"
# lText_FileName = "D:\\Development\\Python\\XLStoCSV\\ExcelSource\\DailyTransactions2019829782539a7-8520-4185-a36a-9b3e415ae22f.csv"

# sys.argv[0] turns out to be the python source program's name, and not the 1st parameter
if len(sys.argv) >= 3:
    # print(len(sys.argv))
    lExcel_FileName = str(sys.argv[1])
    # print(lExcel_FileName)
    lText_FileName = str(sys.argv[2])
    # print(lText_FileName)
    lSeparator = ","

    # create a new text file for writing to
    lOutFile = open(lText_FileName, "w+")

    workbook = xlrd.open_workbook(lExcel_FileName)
    sheet = workbook.sheet_by_index(0)
    lHeaderFound = False
    lHeaderText = "Id,Type,,Result,Code,Account Number,Merchant Number,Merchant Name,Product Code,"

    for row in range(sheet.nrows):
        lRowStr = ""
        for column in range(sheet.ncols):
            lRowStr = lRowStr + str(sheet.cell_value(row, column)) + lSeparator
        if lHeaderFound != True:
            lHeaderLen = len(lHeaderText)
            if lRowStr[:lHeaderLen] == lHeaderText:  # lRowStr[0:lHeaderLen]
                lHeaderFound = True
        if lHeaderFound:
            lOutFile.write(lRowStr + "\n")

    lOutFile.close()
else:
    print("too few arguments")
