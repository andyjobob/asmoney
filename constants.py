
########################################
# Constants for journal transaction table
########################################
COL_DATE = 1
COL_RFNO = 2
COL_DRMRCH = 3
COL_DRATYP = 4
COL_DRACNT = 5
COL_DRASUB = 6
COL_AMNT = 7
COL_CRMRCH = 8
COL_CRATYP = 9
COL_CRACNT = 10
COL_CRASUB = 11
COL_DESC = 12
COL_DUPL = 13
COL_DSRC = 14
COL_SPLIT = 15

journal_header_list = ["Date", "RefNo",
                       "Dr-Merchant", "Dr-Account-Type", "Dr-Account", "Dr-Sub-Account",
                       "Amount",
                       "Cr-Merchant", "Cr-Account-Type", "Cr-Account", "Cr-Sub-Account",
                       "Description", "Duplicate", "Source", "SplitNo"]

########################################
# Constants for raw transaction table
########################################
COL_RAW_RFNO = 1
COL_RAW_DATE = 2
COL_RAW_AMNT = 3
COL_RAW_DESC = 4

raw_header_list = ["RefNo", "Date", "Amount", "Description"]

########################################
# Constants for keyword lookup table
########################################
keyword_header_list = ["Keyword",
                       "Dr-Merchant", "Dr-Account-Type", "Dr-Account", "Dr-Sub-Account",
                       "Cr-Merchant", "Cr-Account-Type", "Cr-Account", "Cr-Sub-Account"]

########################################
# Constants for output transaction table
########################################
COL_OUT_DATE = 1
COL_OUT_RFNO = 2
COL_OUT_MRCH = 3
COL_OUT_ATYP = 4
COL_OUT_ACNT = 5
COL_OUT_ASUB = 6
COL_OUT_AMNT = 7
COL_OUT_DESC = 8

