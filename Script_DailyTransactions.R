library('gmailr')
library('filesstrings')
library('readxl')
library('readr')
library("writexl")
library("dplyr")
library("plyr")

monthlyFilesDir = "//sunfile01/Shared/Fresco/Sales/Sales Files/All Sales (Store, Item, Trans, Cost)/"
dailyFileDir_Archive = "//sunfile01/Shared/Fresco/Sales/Sales Files/Archive/"
dailyFileDir_Working = "//sunfile01/Shared/Fresco/DB/ETL/Daily Transactions/data/"

dailyFileDate = format(as.Date(Sys.time())-1, "%Y-%m-%d")
dailyFileName = "Sun Holdings Daily Transaction Detail.xlsx"
dailyFileFullPath = paste0(dailyFileDir_Working, dailyFileDate, " ", dailyFileName)

mailAccount = "efelix@frescoinvestments.net"
mailSubject = "Sun Holdings Daily Transaction Detail"
mailCredentials = "//sunfile01/Shared/Fresco/DB/ETL/Daily Transactions/credentials.json"
gm_auth_configure(path = mailCredentials)
gm_auth(email = mailAccount)

# EXTRACTS LIST OF MESSAGES BASED ON A GIVEN SUBJECT AND LOOPS THROUGH THE 1ST MESSAGE (MOST RECENT) TO DOWNLOAD ATTACHMENTS
mssgs = gm_messages(search = mailSubject,
                    num_results = NULL,
                    label_ids = NULL,
                    include_spam_trash = NULL,
                    page_token = NULL, 
                    user_id = "me")

for (i in 1:1){
  ids = gm_id(mssgs)
  Mn = gm_message(ids[i], user_id = "me")
  path = dailyFileDir_Archive
  gm_save_attachments( Mn, attachment_id = NULL, path, user_id = "me")
}

# RENAMES FILE DOWNLOADED AND GIVES IT A TIMESTAMP
file.rename(paste0(dailyFileDir_Archive, dailyFileName), 
            paste0(dailyFileDir_Archive, dailyFileDate," ", dailyFileName)
)

file.copy(
  paste0(dailyFileDir_Archive, dailyFileDate," ", dailyFileName),
  paste0(dailyFileDir_Working, dailyFileDate," ", dailyFileName),
  overwrite = TRUE
)

# CREATION OF DATAFRAME TO CAPTURE CURRENT DAILY TRANSACTIONS FILE TO WORK WITH
dailyFile = read_excel(
  dailyFileFullPath,
  col_types = c("text", "numeric", "numeric", "text", "numeric", "text", "numeric", "text", "numeric", "numeric", "numeric"),
  skip = 3
)

dailyFile$Date <- as.Date(dailyFile$Date, format = "%m/%d/%y")
colnames(dailyFile)[8] <- "Description"

# CREATION OF DATAFRAMES TO CAPTURE MONTH OF MAX DATE IN DAILY FILE AND MONTH OF MIN DATE IN DAILY FILE
MaxDateInFile = max(dailyFile$Date)
MaxDay = format(MaxDateInFile, "%d")
MaxMonth = format(MaxDateInFile, "%m")
MaxYear = format(MaxDateInFile, "%Y")
salesFileMax = paste0(monthlyFilesDir ,MaxYear,"/",MaxYear, " - ", MaxMonth, " Transactions and Cost.xlsx")

MinDateInFile = min(dailyFile$Date)
MinDay = format(MinDateInFile, "%d")
MinMonth = format(MinDateInFile, "%m")
MinYear = format(MinDateInFile, "%Y")
salesFileMin = paste0(monthlyFilesDir, MinYear, "/", MinYear, " - ", MinMonth, " Transactions and Cost.xlsx")


if (salesFileMax == salesFileMin) {
  
    print("THERE IS ONLY 1 MONTH IN DAILY TRANSACTIONS FILE")
    
    salesStaging_0 = read_excel(
      salesFileMax,
      col_types = c("date", "numeric", "numeric", "text", "numeric", "text", "numeric", "text", "numeric", "numeric", "numeric")
    )
    
    # CREATION OF STAGING TABLE WITH SALES DATES TO WORK WITH
    salesStaging <- rbind(salesStaging_0)
    salesStaging <- salesStaging[salesStaging$Date < format(MinDateInFile-1, "%Y-%m-%d"),]
    salesStaging <- rbind(salesStaging, dailyFile)
    
    # CREATION OF RESULTING DATAFRAMES WITH MAX MONTH [RESULTING_0]
    salesResulting_0 <- subset(salesStaging,
                               Date >= format(as.Date(paste0(MaxYear, "-", MaxMonth, "-", "01"), "%Y-%m-%d") - 1, "%Y-%m-%d") & 
                                 Date <= format(as.Date(paste0(MaxYear, "-", MaxMonth, "-", MaxDay), "%Y-%m-%d") - 0, "%Y-%m-%d"))
    
    salesResulting_0 <- salesResulting_0[order(salesResulting_0$Date),]
  
    # CREATION OF DIRECTORY AND FILE WITH MAX MONTH SALES FILE
    dir.create(paste0(monthlyFilesDir, MaxYear, "/"))
    write_xlsx(list("Sales Transactions" = salesResulting_0),
               path = salesFileMax,
               col_names = TRUE, format_headers = TRUE)
  
} else {
  
    print("THERE ARE 2 MONTHS IN DAILY TRANSACTIONS FILE")
  
  if (MaxDay == "01") {
    
      print("1ST DAY OF MAX MONTH")
      
      salesStaging_1 = read_excel(
        salesFileMin,
        col_types = c("date", "numeric", "numeric", "text", "numeric", "text", "numeric", "text", "numeric", "numeric", "numeric")
      )
      
      # CREATION OF STAGING TABLE WITH SALES DATES TO WORK WITH
      salesStaging <- rbind(salesStaging_1)
      salesStaging <- salesStaging[salesStaging$Date < format(MinDateInFile-1, "%Y-%m-%d"),]
      salesStaging <- rbind(salesStaging, dailyFile)
      
      # CREATION OF RESULTING DATAFRAMES WITH MAX MONTH [RESULTING_0] AND MIN MONTH [RESULTING_1]
      salesResulting_0 <- subset(salesStaging,
                                 Date >= format(as.Date(paste0(MaxYear, "-", MaxMonth, "-", "01"), "%Y-%m-%d") - 1, "%Y-%m-%d") & 
                                   Date <= format(as.Date(paste0(MaxYear, "-", MaxMonth, "-", MaxDay), "%Y-%m-%d") - 0, "%Y-%m-%d"))
      
      salesResulting_1 <- subset(salesStaging,
                                 Date >= format(as.Date(paste0(MinYear, "-", MinMonth, "-", "01"), "%Y-%m-%d") - 1, "%Y-%m-%d") & 
                                   Date <= format(as.Date(paste0(MaxYear, "-", MaxMonth, "-", "01"), "%Y-%m-%d") - 1, "%Y-%m-%d"))
      
      salesResulting_0 <- salesResulting_0[order(salesResulting_0$Date),]
      salesResulting_1 <- salesResulting_1[order(salesResulting_1$Date),]
      
      # CREATION OF DIRECTORY AND FILE WITH MAX MONTH SALES FILE AND MIN MONTH SALES FILE
      dir.create(paste0(monthlyFilesDir, MaxYear, "/"))
      write_xlsx(list("Sales Transactions" = salesResulting_0),
                 path = salesFileMax,
                 col_names = TRUE, format_headers = TRUE)
      
      dir.create(paste0(monthlyFilesDir, MinYear, "/"))
      write_xlsx(list("Sales Transactions" = salesResulting_1),
                 path = salesFileMin,
                 col_names = TRUE, format_headers = TRUE)
    
  } else {
    
      print("NOT 1ST DAY OF MAX MONTH")
      
      salesStaging_0 = read_excel(
        salesFileMax,
        col_types = c("date", "numeric", "numeric", "text", "numeric", "text", "numeric", "text", "numeric", "numeric", "numeric")
      )
      
      salesStaging_1 = read_excel(
        salesFileMin,
        col_types = c("date", "numeric", "numeric", "text", "numeric", "text", "numeric", "text", "numeric", "numeric", "numeric")
      )
      
      # CREATION OF STAGING TABLE WITH SALES DATES TO WORK WITH
      salesStaging <- rbind(salesStaging_0, salesStaging_1)
      salesStaging <- salesStaging[salesStaging$Date < format(MinDateInFile-1, "%Y-%m-%d"),]
      salesStaging <- rbind(salesStaging, dailyFile)
      
      # CREATION OF RESULTING DATAFRAMES WITH MAX MONTH [RESULTING_0] AND MIN MONTH [RESULTING_1]
      salesResulting_0 <- subset(salesStaging,
                                 Date >= format(as.Date(paste0(MaxYear, "-", MaxMonth, "-", "01"), "%Y-%m-%d") - 1, "%Y-%m-%d") & 
                                   Date <= format(as.Date(paste0(MaxYear, "-", MaxMonth, "-", MaxDay), "%Y-%m-%d") - 0, "%Y-%m-%d"))
      
      salesResulting_1 <- subset(salesStaging,
                                 Date >= format(as.Date(paste0(MinYear, "-", MinMonth, "-", "01"), "%Y-%m-%d") - 1, "%Y-%m-%d") & 
                                   Date <= format(as.Date(paste0(MaxYear, "-", MaxMonth, "-", "01"), "%Y-%m-%d") - 1, "%Y-%m-%d"))
      
      salesResulting_0 <- salesResulting_0[order(salesResulting_0$Date),]
      salesResulting_1 <- salesResulting_1[order(salesResulting_1$Date),]
      
      # CREATION OF DIRECTORY AND FILE WITH MAX MONTH SALES FILE AND MIN MONTH SALES FILE
      dir.create(paste0(monthlyFilesDir, MaxYear, "/"))
      write_xlsx(list("Sales Transactions" = salesResulting_0),
                 path = salesFileMax,
                 col_names = TRUE, format_headers = TRUE)
      
      dir.create(paste0(monthlyFilesDir, MinYear, "/"))
      write_xlsx(list("Sales Transactions" = salesResulting_1),
                 path = salesFileMin,
                 col_names = TRUE, format_headers = TRUE)
    
  }
  
}

# DELETE DAILY FILE DOWNLOADED IN WORKING DIRECTORY
unlink(dailyFileFullPath)

print("EXECUTION COMPLETED....")

q("no", 0, FALSE)
