import sys
import os
import re
import openpyxl

################################################################################
# Description:
#
#  Arguments:
#    1. dir      - Directory where transactions are stored. Make sure dir only
#                  contains transactions (as *.xlsx) and no sub directories
#    2. out_file - File name to write output data to
#
# ex. > python nano_fram_bfm.py ./dlog_dir out_file.csv
#
# Revision:
#   0.1, Andy Sciullo, 09/01/2016, Initial Version
#
################################################################################
################################################################################
# Notes:
#   - Files must be named "*_Transactions.xlsx"
#   - Workbooks should have all transactions in a sheet named "Journal"
#
################################################################################

################################################################################
### Constants
################################################################################
COL_DATE = 1
COL_RFNO = 2
COL_MRCH = 3
COL_ATYP = 4
COL_ACNT = 5
COL_ASUB = 6
COL_AMNT = 7
COL_DESC = 8
COL_DUPL = 9

################################################################################
### Global Variables
################################################################################
dictDrDate = dict()
dictDrMrch = dict()
dictDrAtyp = dict()
dictDrAcnt = dict()
dictDrAsub = dict()
dictDrAmnt = dict()
dictDrDesc = dict()
dictCrDate = dict()
dictCrMrch = dict()
dictCrAtyp = dict()
dictCrAcnt = dict()
dictCrAsub = dict()
dictCrAmnt = dict()
dictCrDesc = dict()
arrDrRefNo = []
arrCrRefNo = []

################################################################################
# Function process_workbook
#
# Description:
#   Check the given file name for correct naming format and read contents for
#   BFM data.
#
#  Arguments:
#    1. file_name - Name of file to process.
#
#  Return Values:
#    None
#
################################################################################
def process_workbook(wb_name):
    # Init variables

    # Search for files with correct name format
    print "  Processing Workbook: ", wb_name
    wb_temp = openpyxl.load_workbook(os.path.join(dir, wb_name))
    sheet = wb_temp.get_sheet_by_name('Journal')
    for rowNum in range(2, sheet.max_row + 1):  # skip the first row
        refNo = sheet.cell(row=rowNum, column=COL_RFNO).value
        amnt = sheet.cell(row=rowNum, column=COL_AMNT).value
        dupl = sheet.cell(row=rowNum, column=COL_DUPL).value
        if not dupl:
            if amnt > 0:
                if refNo in arrDrRefNo:
                    print "    Duplicate reference number found on Debit side -  RFNO: ", refNo, ", AMNT: ", amnt
                else:
                    arrDrRefNo.append(refNo)
                dictDrDate[refNo] = sheet.cell(row=rowNum, column=COL_DATE).value
                dictDrMrch[refNo] = sheet.cell(row=rowNum, column=COL_MRCH).value
                dictDrAtyp[refNo] = sheet.cell(row=rowNum, column=COL_ATYP).value
                dictDrAcnt[refNo] = sheet.cell(row=rowNum, column=COL_ACNT).value
                dictDrAsub[refNo] = sheet.cell(row=rowNum, column=COL_ASUB).value
                dictDrAmnt[refNo] = sheet.cell(row=rowNum, column=COL_AMNT).value
                dictDrDesc[refNo] = sheet.cell(row=rowNum, column=COL_DESC).value
            else:
                if refNo in arrCrRefNo:
                    print "    Duplicate reference number found on Credit side -  RFNO: ", refNo, ", AMNT: ", amnt
                else:
                    arrCrRefNo.append(refNo)
                dictCrDate[refNo] = sheet.cell(row=rowNum, column=COL_DATE).value
                dictCrMrch[refNo] = sheet.cell(row=rowNum, column=COL_MRCH).value
                dictCrAtyp[refNo] = sheet.cell(row=rowNum, column=COL_ATYP).value
                dictCrAcnt[refNo] = sheet.cell(row=rowNum, column=COL_ACNT).value
                dictCrAsub[refNo] = sheet.cell(row=rowNum, column=COL_ASUB).value
                dictCrAmnt[refNo] = sheet.cell(row=rowNum, column=COL_AMNT).value
                dictCrDesc[refNo] = sheet.cell(row=rowNum, column=COL_DESC).value
        #else:
        #    print "    Excluding marked duplicate - RFNO: ", refNo, ", AMNT: ", amnt

################################################################################
# Function print_transactions
#
# Description:
#   Check the given file name for correct naming format and read contents for
#   BFM data.
#
#  Arguments:
#    1. file_name - Name of file to process.
#
#  Return Values:
#    None
#
################################################################################
def print_transactions(wb_name):
    print "\n  Writing Output: ", wb_name

    wb_out = openpyxl.load_workbook(wb_name)
    sheet = wb_out.get_sheet_by_name('Journal')

    # Clear out previous journal data
    for rowNum in range(2, sheet.max_row + 1):  # skip the first row
        for colNum in range(1, sheet.max_column + 1):
            sheet.cell(row=rowNum, column=colNum).value = ""

    for refNo in arrDrRefNo:
        if refNo not in arrCrRefNo:
            arrDrRefNo.remove(refNo)
            print "    Debit reference number not found for Credit: ", refNo

    for refNo in arrCrRefNo:
        if refNo not in arrCrRefNo:
            arrCrRefNo.remove(refNo)
            print "    Credit reference number not found for Debit: ", refNo

    rowCnt = 2
    for refNo in arrDrRefNo:
        #print("%s,%s,%s,%s,%s,%s,%s,%s" % (refNo, dictDrDate[refNo], dictDrMrch[refNo], dictDrAtyp[refNo], dictDrAcnt[refNo], dictDrAsub[refNo], \
        #    dictDrAmnt[refNo], dictDrDesc[refNo]))
        #print("%s,%s,%s,%s,%s,%s,%s,%s" % (refNo, dictCrDate[refNo], dictCrMrch[refNo], dictCrAtyp[refNo], dictCrAcnt[refNo], dictCrAsub[refNo], \
        #    dictCrAmnt[refNo], dictCrDesc[refNo]))

        # Debit row
        sheet.cell(row=rowCnt, column=COL_DATE).value = dictDrDate[refNo]
        sheet.cell(row=rowCnt, column=COL_RFNO).value = refNo
        sheet.cell(row=rowCnt, column=COL_MRCH).value = dictDrMrch[refNo]
        sheet.cell(row=rowCnt, column=COL_ATYP).value = dictDrAtyp[refNo]
        sheet.cell(row=rowCnt, column=COL_ACNT).value = dictDrAcnt[refNo]
        sheet.cell(row=rowCnt, column=COL_ASUB).value = dictDrAsub[refNo]
        sheet.cell(row=rowCnt, column=COL_AMNT).value = dictDrAmnt[refNo]
        sheet.cell(row=rowCnt, column=COL_DESC).value = dictDrDesc[refNo]
        rowCnt += 1

        # Credit row
        sheet.cell(row=rowCnt, column=COL_DATE).value = dictCrDate[refNo]
        sheet.cell(row=rowCnt, column=COL_RFNO).value = refNo
        sheet.cell(row=rowCnt, column=COL_MRCH).value = dictCrMrch[refNo]
        sheet.cell(row=rowCnt, column=COL_ATYP).value = dictCrAtyp[refNo]
        sheet.cell(row=rowCnt, column=COL_ACNT).value = dictCrAcnt[refNo]
        sheet.cell(row=rowCnt, column=COL_ASUB).value = dictCrAsub[refNo]
        sheet.cell(row=rowCnt, column=COL_AMNT).value = dictCrAmnt[refNo]
        sheet.cell(row=rowCnt, column=COL_DESC).value = dictCrDesc[refNo]
        rowCnt += 1

    sheet.freeze_panes = 'A2'
    wb_out.save(wb_name)


################################################################################
# Main
#
# Description:
#   Check the given file name for correct naming format and read contents for
#   BFM data.
#
#  Arguments:
#    1. file_name - Name of file to process.
#
#  Return Values:
#    None
#
################################################################################
# Process input arguments
if (len(sys.argv) != 3):
    print("Usage: python asmoney.py <PATH_TO_DIR> <OUTFILE>")
    sys.exit()

dir = str(sys.argv[1])
out_file = str(sys.argv[2])

print "Transaction Dir : ", dir
print "Output File     : ", out_file

# Loop through each file in directory
for file_name in os.listdir(dir):
    if os.path.isfile(os.path.join(dir, file_name)):
        search_obj = re.search(r'.*_Transactions\.xlsx', file_name)
        if search_obj:
            process_workbook(file_name)

print_transactions(out_file)
print "\n ==> ...Ending"








