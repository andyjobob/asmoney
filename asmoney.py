import sys
import os
import re
import openpyxl

# Files must be named "*_Transactions.xlsx"
# Workbooks should have all transactions in a sheet named "Journal"


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
    search_obj = re.search(r'.*_Transactions\.xlsx', wb_name)
    if search_obj:
        print "  Found Workbook: ", wb_name
        #wb_temp = openpyxl.load_workbook(os.path.join(dir, wb_name))
        #sheet = wb_temp.get_sheet_by_name('Journal')

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
#num_files = sum(os.path.isfile(os.path.join(dir, f)) for f in os.listdir(dir))
#print "         Found Files: ", num_files
for file_name in os.listdir(dir):
    if os.path.isfile(os.path.join(dir, file_name)):
        process_workbook(file_name)
        #sys.stdout.write("\r         Files Processed: %d" % file_count)

# Close output file and report number of files processed
#out_handle.close()
#print "     Files processed: ", file_count
print "\n ==> ...Ending"


#os.getcwd()
#os.chdir()




#sheet = wb.get_sheet_by_name('Sheet1')

#max_column = sheet.max_column
#max_row = sheet.max_row

#for cellObj in sheet.columns[1]:
#    print(cellObj.value)

#for rowNum in range(2, sheet.max_row):  # skip the first row


#wb.save('example.xlsx')

#sheet['A1'] = 'Hello World!'

#sheet.freeze_panes = 'A2'

# Search input directory for files *_Transactions.xlsx


# Open each workbook and select the "Journal" tab

# Open out file in
#wb_out = openpyxl.load_workbook(out_file)







