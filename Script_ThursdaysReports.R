# 0.1- LIBRARIES TO USE--------

# install.packages('gmailr')
# install.packages('filesstrings')
# install.packages('readxl')
# install.packages('readr')
# install.packages("writexl")

library('gmailr')
library('filesstrings')
library('readxl')
library('readr')
library("writexl")

# 0.2- DECLARATIONS AND GMAIL CONFIG--------

# OBTAIN CURRENT DIRECTORY
print("OBTAINING WORKING DIRECTORY ......")
displayWD = "//sunfile01/Shared/Fresco/DB/ETL/Thursday Reports/"

# Gmail Account Credentials
mailCredentials = paste0(displayWD,"credentials.json")
mailAccount = "efelix@frescoinvestments.net"
#mailQuery = "from:(Lynnette-Howell@gnc-hq.com) subject:(Thursday Reports) has:attachment"
mailQuery = "from:(gnc-hq.com) subject:(Thursday Reports) has:attachment"

    # CONFIGURING MAIL CONNECTION WITH GMAIL CREDENTIALS
    print("READING LOGIN CREDENTIALS ......")
    gm_auth_configure(path = mailCredentials)
    
    # AUTHENTICATING WITH PRE-AUTHORIzED ACCOUNT
    print("AUTHENTICATING WITH GMAIL ......")
    gm_auth(email = mailAccount)
    
    # EXTRACTS LIST OF MESSAGES BASED ON A GIVEN GMAIL QUERY
    print("READING EMAILS ......")
    print("OBTAINING RECIEVED DATE ......")
    mssgs = gm_messages(search = mailQuery,
                        num_results = NULL,
                        label_ids = NULL,
                        include_spam_trash = NULL,
                        page_token = NULL, 
                        user_id = "me")
    
    ids = gm_id(mssgs)
    Mn = gm_message(ids[1], user_id = "me")
    
    # TRANSFORMING DATE TO YYYY-MM-DD FORMAT
    dateGMAIL = substr(gm_date(Mn), 1, str_length(gm_date(Mn))-6)
    receivedDate = as.Date(dateGMAIL, "%a, %d %b %Y %H:%M:%S")
    
    receivedDate = paste0(receivedDate)
    
    datestamps = read.csv(file = paste0(displayWD, "datestamps.csv"))
    names(datestamps)[1] = "DateReceived"
    
    receivedDate = data.frame("DateReceived"=receivedDate)
    print(paste0("Email Received on: ", receivedDate[1,1]))
    
    # OBTAINING FROM AND TO DATES OF REPORTS
    tsDF = merge(receivedDate, datestamps, by.x = "DateReceived")
    tsFrom = paste0(tsDF[1,3])
    tsTo = paste0(tsDF[1,4])
    tsFull = paste0(tsFrom, "to", tsTo, " ")

# Attachements Download Folder
filesPath = paste0(displayWD, "temporary")

# List of Destination Folders (Archive Folders)
folderADAS = "S:/Fresco/Sales/Sales Files/ADAS/Weekly Files"
folderConcealedDamage = "S:/Fresco/Shipments/Concealed Damage/Weekly Files"
folderCoupon = "S:/Fresco/Coupon/Coupon Weekly files/Coupon Detail"
folderDiscrepancies = "S:/Fresco/Shipments/Discrepancies"
folderGiftCards = "S:/Fresco/Gift Cards/Weekly Files"
folderGNCDelivers = "S:/Fresco/Shipments/GNC Delivers/GNCDELIV Weekly Files"
folderInTransit = "S:/Fresco/Shipments/In Transit/InTransit Weekly Files"
folderInventoryOnHands = "S:/Fresco/Inventory/Inventory On Hand Weekly Files"
folderKnownLoss = "S:/Fresco/Known Loss/KL Weekly Files"
folderLeadComp = "S:/Fresco/Products List/Lead-Component/Weekly Files"
folderProduct = "S:/Fresco/Shipments/Product/Product Ship Weekly Files"
folderRecall = "S:/Fresco/Recall/Recall Weekly Files"
folderReceivedDates = "S:/Fresco/Shipments/Received Dates/Weekly Files"
folderReturns = "S:/Fresco/Sales/Returns/Weekly Files"
folderSupplies = "S:/Fresco/Shipments/Supplies/Supplies Weekly Files"
folderTransfers = "S:/Fresco/Shipments/Transfers/Transfers Weekly NF"

# List of Destination Folders (for Importing to SQL)
SQLmainFolder = paste0(displayWD, "data")
SQLfolderADAS = paste0(SQLmainFolder, "/ADAS")
SQLfolderConcealedDamage = paste0(SQLmainFolder, "/Concealed Damage")
SQLfolderCoupon = paste0(SQLmainFolder, "/Coupon")
SQLfolderDiscrepancies = paste0(SQLmainFolder, "/Discrepancies")
SQLfolderGiftCards = paste0(SQLmainFolder, "/Gift Cards")
SQLfolderGNCDelivers = paste0(SQLmainFolder, "/GNC Delivers")
SQLfolderInTransit = paste0(SQLmainFolder, "/In Transit")
SQLfolderInventoryOnHands = paste0(SQLmainFolder, "/Inventory On Hand")
SQLfolderKnownLoss = paste0(SQLmainFolder, "/Known Loss")
SQLfolderLeadComp = paste0(SQLmainFolder, "/Lead-Component")
SQLfolderProduct = paste0(SQLmainFolder, "/Product")
SQLfolderRecall = paste0(SQLmainFolder, "/Recall")
SQLfolderReceivedDates = paste0(SQLmainFolder, "/Received Dates")
SQLfolderReturns = paste0(SQLmainFolder, "/Returns")
SQLfolderSupplies = paste0(SQLmainFolder, "/Supplies")
SQLfolderTransfers = paste0(SQLmainFolder, "/Transfers")

# List of Destination FileNames
fileNameADAS = "ADAS Report (email).xlsx"
fileNameConcealedDamage = "Reported Date (email).xlsx"
fileNameCoupon = "Coupon Redemption Details (email).xlsx"
fileNameDiscrepancies = "Discrepancies Reported and Closed (email).xlsx"
fileNameGiftCards = "Gift Cards (email).xlsx"
fileNameGNCDelivers = "GNC Delivers Ships (email).xlsx"
fileNameInTransit = "In Transit (email).xlsx"
fileNameInventoryOnHands = "Store On Hands (email).xlsx"
fileNameKnownLoss = "Known Loss (email).xlsx"
fileNameLeadComp = "Lead-Component List (email).xlsx"
fileNameProduct = "Wholesale Ships (email).xlsx"
fileNameRecall = "Recalls (email).xlsx"
fileNameReceivedDates = "Rec Invoices (email).xlsx"
fileNameReturns = "Returns (email).xlsx"
fileNameSupplies = "Supply Ships (email).xlsx"
fileNameTransfers = "Transfers (email).xlsx"

# 1.1- EMAIL ATTACHMENTS' DOWNLOAD--------

# READING EMAILS WITH SPECIFIED CRITERIA
mssgs = gm_messages(search = mailQuery,
                    num_results = NULL,
                    label_ids = NULL,
                    include_spam_trash = NULL,
                    page_token = NULL, 
                    user_id = "me")

# GOES THROUGH THE FIRST MESSAGE (MOST RECENT ONE) TO DOWNLOAD ATTACHMENTS
print("DOWNLOADING FILES TO TEMPORARY FOLDER ......")
for (i in 1:1){
  ids = gm_id(mssgs)
  Mn = gm_message(ids[i], user_id = "me")
  path = filesPath
  gm_save_attachments( Mn, attachment_id = NULL, path, user_id = "me")
}


# 2.1- SPLITTING FILES INTO XLSX FILES--------
print("MOVING FILES TO RESPECTIVE FOLDERS .....")

# GENERATE LIST OF FILES IN DOWNLOADS FOLDER
filesArray0 = list.files(path = filesPath)
filesArray = toupper(filesArray0)

print(filesArray)

# GOES THROUGH EACH FILE TO OPEN IT AND SPLIT WORKSHEETS
for (i in 1:length(filesArray)){
  if ( grepl("WHOLESALE SHIPS", filesArray[i])) {

        excelFile = paste0(filesPath, "/", filesArray0[i])
        
        #*********1.1 - SPLITTING PRODUCT TAB***************
        print("SPLITTING PRODUCT TAB")
        folderAssgined = folderProduct
        SQLfolderAssgined = SQLfolderProduct
        fileNameAssigned = fileNameProduct

        skipAssgined = 4
        numCols = 9
        
        excelDF = read_excel(excelFile,
                             sheet = "daily gross w'sale",
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                             col_names = c("CENTER_NBR", "ITEM_NBR", "PRODUCT_DESC", "SIZE", "UNITS SHIPPED", "NET COST SHIPPED", "SHIP_DATE", "INVOICE NBR", "CONV PURCHASE"), 
                             col_types = c("numeric", "numeric", "text", "numeric", "numeric", "numeric", "date", "numeric", "text"), 
                             na = "", 
                             skip = skipAssgined
                             )
        
        write_xlsx(list("product" = excelDF), 
                   path = paste0(folderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("product" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        
        #*********1.2 - SPLITTING SUPPLIES TAB***************
        print("SPLITTING SUPPLIES TAB")
        folderAssgined = folderSupplies
        SQLfolderAssgined = SQLfolderSupplies
        fileNameAssigned = fileNameSupplies
        
        skipAssgined = 3
        numCols = 8
        
        excelDF = read_excel(excelFile,
                             sheet = "daily gross w'sale supplies",
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                             col_names = TRUE,
                             col_types = c("numeric", "numeric", "text", "numeric", "numeric", "numeric", "date", "numeric"),
                             na = "",
                             skip = skipAssgined
                             )
        
        write_xlsx(list("supplies" = excelDF), 
                   path = paste0(folderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("supplies" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        #*********1.3 - SPLITTING GNC DELIVERS TAB***************
        print("SPLITTING GNC DELIVERS TAB")
        folderAssgined = folderGNCDelivers
        SQLfolderAssgined = SQLfolderGNCDelivers
        fileNameAssigned = fileNameGNCDelivers
        
        skipAssgined = 3
        numCols = 8
        
        excelDF = read_excel(excelFile,
                             sheet = "daily gross GNC delivers",
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                             col_names = TRUE,
                             col_types = c("numeric", "numeric", "text", "numeric", "numeric", "numeric", "date", "numeric"),
                             na = "",
                             skip = skipAssgined
                             )
        
        write_xlsx(list("daily gross GNC delivers" = excelDF), 
                   path = paste0(folderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("daily gross GNC delivers" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        #*********1.4 and 1.5 - SPLITTING INVOICES TABS*************
        print("SPLITTING INVOICES TABS")
        folderAssgined = folderReceivedDates
        SQLfolderAssgined = SQLfolderReceivedDates
        fileNameAssigned = fileNameReceivedDates
        
        skipAssgined = 3
        numCols = 4
        
        excelDF_RD1 = read_excel(excelFile,
                                 sheet = "w'sale invoices",
                                 range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                                 col_names = TRUE,
                                 col_types = c("numeric", "numeric", "date", "date"),
                                 na = "",
                                 skip = skipAssgined
                                 )
        
        excelDF_RD2 = read_excel(excelFile,
                                 sheet = "w'sale invoices rec",
                                 range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                                 col_names = TRUE,
                                 col_types = c("numeric", "numeric", "date", "date"),
                                 na = "",
                                 skip = skipAssgined
                                 )
        
        write_xlsx(list("w'sale invoices" = excelDF_RD1, "w'sale invoices rec" = excelDF_RD2), 
                   path = paste0(folderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("w'sale invoices" = excelDF_RD1, "w'sale invoices rec" = excelDF_RD2), 
                   path = paste0(SQLfolderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        #*********1.6, 1.7 and 1.8 - SPLITTING TRANSFER TABS*************
        print("SPLITTING TRANSFER TABS")
        folderAssgined = folderTransfers
        SQLfolderAssgined = SQLfolderTransfers
        fileNameAssigned = fileNameTransfers
        
        skipAssgined = 2
        numCols1 = 11
        numCols2 = 12
        numCols3 = 12
        
        excelDF_T1 = read_excel(excelFile,
                                sheet = "transfers - send",
                                range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols1)),
                                col_names = TRUE,
                                col_types = c("numeric", "numeric", "date", "date", "numeric", "text", "numeric", "numeric", "numeric", "numeric", "text"),
                                na = "",
                                skip = skipAssgined
                                )
        
        excelDF_T2 = read_excel(excelFile,
                                sheet = "transfers - rec",
                                range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols2)),
                                col_names = TRUE,
                                col_types = c("numeric", "numeric", "date", "date", "numeric", "text", "numeric", "numeric", "numeric", "numeric", "numeric", "text"),
                                na = "",
                                skip = skipAssgined
                                )
        
        excelDF_T3 = read_excel(excelFile,
                                sheet = "transfers - send (other)",
                                range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols3)),
                                col_names = TRUE,
                                col_types = c("numeric", "numeric", "date", "date", "numeric", "text", "numeric", "numeric", "numeric", "numeric", "numeric", "text"),
                                na = "",
                                skip = skipAssgined
                                )
        
        write_xlsx(list("transfers - send" = excelDF_T1, "transfers - rec" = excelDF_T2, "transfers - send (other)" = excelDF_T3), 
                   path = paste0(folderAssgined, "/", paste0(tsFrom, " to ", tsTo, " "), fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("transfers - send" = excelDF_T1, "transfers - rec" = excelDF_T2, "transfers - send (other)" = excelDF_T3), 
                   path = paste0(SQLfolderAssgined, "/", paste0(tsFrom, " to ", tsTo, " "), fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        #*********1.9 - SPLITTING IN TRANSITS TAB***************
        print("SPLITTING IN TRANSITS TAB")
        folderAssgined = folderInTransit
        SQLfolderAssgined = SQLfolderInTransit
        fileNameAssigned = fileNameInTransit
        
        skipAssgined = 1
        numCols = 8
        
        excelDF = read_excel(excelFile,
                             sheet = "in transits",
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                             col_names = TRUE,
                             col_types = c("numeric", "numeric", "text", "numeric", "numeric", "numeric", "numeric", "date"),
                             na = "",
                             skip = skipAssgined
                             )
        
        write_xlsx(list("in transits" = excelDF), 
                   path = paste0(folderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("in transits" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        #*********1.10 - SPLITTING IN LEAD-COMP TAB***************
        print("SPLITTING LEAD-COMP TAB")
        folderAssgined = folderLeadComp
        SQLfolderAssgined = SQLfolderLeadComp
        fileNameAssigned = fileNameLeadComp
        
        skipAssgined = 0
        numCols = 3
        
        excelDF = read_excel(excelFile,
                             sheet = "lead - comp",
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                             col_names = TRUE,
                             col_types = c("numeric", "numeric", "numeric"),
                             na = "",
                             skip = skipAssgined
                             )
        
        write_xlsx(list("lead - comp" = excelDF), 
                   path = paste0(folderAssgined, "/", tsTo, " ", fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("lead - comp" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsTo, " ", fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        
        #*********1.11 - FINISHING BY DELETING FILE***************
        unlink(excelFile) #CHANGE PRINT COMMAND FOR UNLINK TO REMOVE FILE IN TEMP FOLDER
        next
    
  } else if ( grepl("KNOWN LOSS", filesArray[i])) {
    
        excelFile = paste0(filesPath, "/", filesArray0[i])
        
        #*********2.1 - SPLITTING KNOWN LOSS TAB***************
        print("SPLITTING KNOWN LOSS TAB")
        folderAssgined = folderKnownLoss
        SQLfolderAssgined = SQLfolderKnownLoss
        fileNameAssigned = fileNameKnownLoss
        
        skipAssgined = 2
        numCols = 11
        
        excelDF = read_excel(excelFile,
                             sheet = "known loss",
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                             col_names = TRUE,
                             col_types = c("numeric", "numeric", "numeric", "text", "numeric", "numeric", "numeric", "date", "date", "text", "numeric"),
                             na = "",
                             skip = skipAssgined
                             )
        
        write_xlsx(list("known loss" = excelDF), 
                   path = paste0(folderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("known loss" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        #*********2.2 - SPLITTING RECALL TAB***************
        print("SPLITTING RECALL TAB")
        folderAssgined = folderRecall
        SQLfolderAssgined = SQLfolderRecall
        fileNameAssigned = fileNameRecall
        
        skipAssgined = 2
        numCols = 12
        
        excelDF = read_excel(excelFile,
                             sheet = "recalls",
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                             col_names = TRUE,
                             col_types = c("numeric", "numeric", "numeric", "text", "numeric", "numeric", "numeric", "date", "date", "text", "numeric", "numeric"),
                             na = "",
                             skip = skipAssgined
                             )
        
        write_xlsx(list("recalls" = excelDF), 
                   path = paste0(folderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("recalls" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        
        #*********2.3 - FINISHING BY DELETING FILE***************
        unlink(excelFile) #CHANGE PRINT COMMAND FOR UNLINK TO REMOVE FILE IN TEMP FOLDER
        next
    
  } else if ( grepl("REPORTED DATE", filesArray[i])) {
    
        excelFile = paste0(filesPath, "/", filesArray0[i])
        
        #*********3.1 - SPLITTING CONCEALED DAMAGE TAB***************
        print("SPLITTING CONCEALED DAMAGE TAB")
        folderAssgined = folderConcealedDamage
        SQLfolderAssgined = SQLfolderConcealedDamage
        fileNameAssigned = fileNameConcealedDamage
        
        skipAssgined = 3
        numCols = 9
        
        excelDF = read_excel(excelFile,
                             sheet = "Concealed Damage",
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                             col_names = TRUE,
                             col_types = c("numeric", "numeric", "date", "numeric", "text", "numeric", "numeric", "date", "numeric"),
                             na = "",
                             skip = skipAssgined
                             )
        
        write_xlsx(list("Concealed Damage" = excelDF), 
                   path = paste0(folderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("Concealed Damage" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        unlink(excelFile) #CHANGE PRINT COMMAND FOR UNLINK TO REMOVE FILE IN TEMP FOLDER
        next
    
  } else if ( grepl("RETURNS", filesArray[i])) {
    
        excelFile = paste0(filesPath, "/", filesArray0[i])
        
        #*********4.1 - SPLITTING RETURNS TAB***************
        print("SPLITTING RETURNS TAB")
        folderAssgined = folderReturns
        SQLfolderAssgined = SQLfolderReturns
        fileNameAssigned = fileNameReturns
        
        skipAssgined = 2
        numCols = 10
        
        excelDF = read_excel(excelFile,
                             sheet = "summary",
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                             col_names = TRUE,
                             col_types = c("numeric", "numeric", "numeric", "text", "numeric", "text", "numeric", "numeric", "numeric", "date"),
                             na = "",
                             skip = skipAssgined
                             )
        
        write_xlsx(list("summary" = excelDF), 
                   path = paste0(folderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("summary" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        unlink(excelFile) #CHANGE PRINT COMMAND FOR UNLINK TO REMOVE FILE IN TEMP FOLDER
        next
    
  } else if ( grepl("STORE ON HANDS", filesArray[i])) {

        excelFile = paste0(filesPath, "/", filesArray0[i])
        
        #*********5.1 - SPLITTING SOH TAB***************
        print("SPLITTING STORE ON HANDS TAB")
        folderAssgined = folderInventoryOnHands
        SQLfolderAssgined = SQLfolderInventoryOnHands
        fileNameAssigned = fileNameInventoryOnHands
        
        skipAssgined = 1
        numCols = 7
        
        excelDF = read_excel(excelFile,
                             sheet = "ON HANDS",
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)),
                             col_names = c("CENTER NBR", "ITEM NBR", "PRODUCT DESC", "SIZE", "ONHAND QTY", "AVERAGE COST", "CURR W'SALE PRICE"),
                             col_types = c("numeric", "numeric", "text", "numeric", "numeric", "numeric", "numeric"),
                             na = "",
                             skip = skipAssgined
                             )
        
        write_xlsx(list("On Hands" = excelDF), 
                   path = paste0(folderAssgined, "/", tsTo, " ", fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("On Hands" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsTo, " ", fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        unlink(excelFile) #CHANGE PRINT COMMAND FOR UNLINK TO REMOVE FILE IN TEMP FOLDER
        next
    
  } else if ( grepl("GIFT CARDS", filesArray[i])) {

        excelFile = paste0(filesPath, "/", filesArray0[i])
        
        #*********6.1 - SPLITTING GIFT CARDS TAB***************
        print("SPLITTING GIFT CARDS TABS")
        folderAssgined = folderGiftCards
        SQLfolderAssgined = SQLfolderGiftCards
        fileNameAssigned = fileNameGiftCards
        skipAssgined = 3
        
        excelDF_RD1 = read_excel(excelFile,
                                 sheet = "Issued",
                                 col_names = TRUE,
                                 col_types = c("numeric", "text", "text", "text", "numeric", "numeric", "numeric", "date", "date", "numeric", "numeric", "text"),
                                 na = "",
                                 skip = skipAssgined)
        
        excelDF_RD2 = read_excel(excelFile,
                                 sheet = "Redeemed",
                                 col_names = TRUE,
                                 col_types = c("numeric", "text", "text", "text", "numeric", "numeric", "numeric", "date", "date", "numeric", "numeric", "text"),
                                 na = "",
                                 skip = skipAssgined)
        
        write_xlsx(list("Issued" = excelDF_RD1, "Redeemed" = excelDF_RD2), 
                   path = paste0(folderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("Issued" = excelDF_RD1, "Redeemed" = excelDF_RD2), 
                   path = paste0(SQLfolderAssgined, "/", tsFull, fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        
        unlink(excelFile) #CHANGE PRINT COMMAND FOR UNLINK TO REMOVE FILE IN TEMP FOLDER
        next
    
  } else if ( grepl("ADAS", filesArray[i])) {

        excelFile = paste0(filesPath, "/", filesArray0[i])
        
        #*********7.1 - SPLITTING ADAS TAB***************
        print("SPLITTING ADAS TAB")
        folderAssgined = folderADAS
        SQLfolderAssgined = SQLfolderADAS
        fileNameAssigned = fileNameADAS
        
        skipAssgined = 4
        numCols = 9
        
        excelDF = read_excel(excelFile,
                             sheet = 1,
                             range = cell_limits(c(1 + skipAssgined, 1), c(NA, numCols)), 
                             col_names = TRUE, 
                             col_types = c("numeric", "date", "numeric", "date", "numeric", "text", "numeric", "numeric", "numeric"), 
                             na = "", 
                             skip = skipAssgined
                             )
        
        excelDF = subset(excelDF, !excelDF$CENTER_NBR=="")
        
        write_xlsx(list("ADAS" = excelDF), 
                   path = paste0(folderAssgined, "/", tsTo, " ", fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("ADAS" = excelDF), 
                   path = paste0(SQLfolderAssgined, "/", tsTo, " ", fileNameAssigned), 
                   col_names = TRUE, format_headers = TRUE)
      
        unlink(excelFile) #CHANGE PRINT COMMAND FOR UNLINK TO REMOVE FILE IN TEMP FOLDER
        next
    
  } else if ( grepl("DISCREPANCIES", filesArray[i])) {
        print("TRANSFERRING DISCREPANCIES")
    
        excelFile = paste0(filesPath, "/", filesArray0[i])
        
        file.copy(
          paste0(filesPath, "/", filesArray0[i]),
          paste0(folderDiscrepancies, "/", filesArray0[i]),
          overwrite = TRUE)
        file.copy(
          paste0(filesPath, "/", filesArray0[i]),
          paste0(SQLfolderDiscrepancies, "/", filesArray0[i]),
          overwrite = TRUE)
        
        unlink(excelFile) #CHANGE PRINT COMMAND FOR UNLINK TO REMOVE FILE IN TEMP FOLDER
        next
    
  } else if ( grepl("COUPON", filesArray[i])) {
  
        excelFile = paste0(filesPath, "/", filesArray0[i])
        
        #*********9.1 - SPLITTING COUPON TAB***************
        print("SPLITTING COUPON TAB")
        folderAssgined = folderCoupon
        SQLfolderAssgined = SQLfolderCoupon
        fileNameAssigned = fileNameCoupon
        skipAssgined = 3

        excelDF = read_excel(excelFile,
                             sheet = "Detailed Coupon Listing",
                             col_names = TRUE,
                             col_types = c("numeric", "date", "numeric", "numeric", "numeric", "numeric", "numeric", "numeric", "text", "date", "date", "numeric", "numeric", "numeric"),
                             na = "",
                             skip = skipAssgined)

        write_xlsx(list("Detailed Coupon Listing" = excelDF),
                   path = paste0(folderAssgined, "/", paste0(tsFrom, " to ", tsTo, " "), fileNameAssigned),
                   col_names = TRUE, format_headers = TRUE)
        write_xlsx(list("Detailed Coupon Listing" = excelDF),
                   path = paste0(SQLfolderAssgined, "/", paste0(tsFrom, " to ", tsTo, " "), fileNameAssigned),
                   col_names = TRUE, format_headers = TRUE)
        
        
        unlink(excelFile) #CHANGE PRINT COMMAND FOR UNLINK TO REMOVE FILE IN TEMP FOLDER
        print("DONE SPLITTING COUPON TAB")
        next
        
  }
        
}

# 3.1- CLOSE EXECUTION--------

# CLOSE SCRIPT
print("EXECUTION COMPLETED .....")

#q("no", 0, FALSE)
