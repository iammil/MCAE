library(lubridate)
library(dplyr)
library(dplyr)
library(RMySQL)
library(readxl)
library(gridExtra)
library(grid)
library(xtable)
library(xlsx)

source("Create Excel.R")

month.now <- month(as.POSIXlt(Sys.Date(), format="%d/%m/%Y"))
month.name <- month.abb[month.now]
Year.Now <- as.numeric(format(as.Date(Sys.time(), format="%d/%m/%Y"),"%Y"))
month.name <- "Oct"
Year.Now <- "2017"

month.now <- 12
  # month(Sys.Date())

Folder <- "C:/Users/GSB/Downloads"

  
File <- paste("Sample_Actual_MCAE_v",month.name,Year.Now,".xlsx",sep="")

  
fullpath <- paste(Folder,File,sep="/")

 
  
  DF.MCAE <- read_excel(fullpath,sheet = 1,col_names = FALSE,na = "",skip = 8)
  findindex <- which(DF.MCAE$X__3 == "This will be the eligible BIN number")-2
  DF.MCAE  <- slice(DF.MCAE ,1:findindex)
  
  MCAE <- select(DF.MCAE,X__2,X__3,X__4,X__5,X__6,X__7,X__8,X__9,X__10,X__11,X__12,X__13,X__14,X__15,X__16,X__17,X__18,X__19,X__20,X__21,X__22,X__23,X__24,X__25,X__26,X__27,X__28,X__29,X__30,X__31,X__32,X__33,X__34,X__35,X__36,X__37,X__38,X__39,X__40,X__41,X__42,X__43,X__44,X__45,X__46,X__47,X__48,X__49,X__50,X__51,X__52,X__53,X__54,X__55)
  MCAE$X__2 <- as.numeric(MCAE$X__2)-2
  MCAE[is.na(MCAE)]="0"
  MCAE$CREATE_DATE <- substr(as.POSIXlt(Sys.time(),tz="Asia/Bangkok", "%Y-%m-%d %H:%M"),1,20)
  
  MCAE.nrow <- nrow(MCAE)
  
  mydb = dbConnect(MySQL()
                   , user='root'
                   , password='123456'
                   , dbname='mastercard_lounge'
                   , host='localhost')
  i <- 1
  for(i in 1:MCAE.nrow) {
      MCAE_SQL <- paste("INSERT INTO mastercard_lounge.mcae (ID, PROCESS_MONTH, BIN, ISSUER, ISSUER_COUNTRY, MC_REGION, ICA, DEAL_TYPE, SOURCE_CODE, MEMBER_NO, CCARD_NO, TITLE, FIRST_NAME, LAST_NAME, MEMBER_STATUS, PAID_TO_DATE, BENEFIT_TRANSACTION_DATE, DATE_PROCESSED, GUESTS, BATCH_NO, REFERENCE, BENEFIT_CODE, BENEFIT, AIRPORT, TERMINAL, INTERNATIONAL_OR_DOMESTIC, CITY, COUNTRY, BENEFIT_TYPE, BENEFIT_CATEGORY, CLIENT_PAYS_MEMBER_EXPERIENCE, INCLUSIVE_MEMBER_EXPERIENCE, COMPLIMENTARY_MEMBER_EXPERIENCE, CARDHOLDER_PAYS_MEMBER_EXPERIENCE, CLIENT_PAYS_GUEST_EXPERIENCE, INCLUSIVE_GUEST_EXPERIENCE, COMPLIMENTARY_GUEST_EXPERIENCE, CARDHOLDER_PAYS_GUEST_EXPERIENCE, BENEFIT_EXPERIENCE_OFFER, CLIENT_MEMBER, INCLUSIVE_MEMBER, CLIENT_GUEST, INCLUSIVE_GUEST, CLIENT_MEMBER_GUEST, INCLUSIVE_MEMBER_GUEST, COMPLIMENTARY_MEMBER_GUEST, CARDHOLDER_VISIT_CURRENCY_MEMBER, CARDHOLDER_VISIT_AMOUNT_MEMBER, CARDHOLDER_VISIT_CURRENCY_GUEST, CARDHOLDER_VISIT_AMOUNT_GUEST, TOTAL_CARDHOLDER_AMOUNT_CURRENCY, TOTAL_CARDHOLDER_AMOUNT, VENDOR_CODE, USER_INVITATION_CODE, CIN, CREATE_DATE) VALUES (NULL,"
                        ,"date_add('1900-01-01',interval ",MCAE$X__2[i]," day)",",'"
                        ,MCAE$X__3[i],"','"
                        ,MCAE$X__4[i],"','"
                        ,MCAE$X__5[i],"','"
                        ,MCAE$X__6[i],"','"
                        ,MCAE$X__7[i],"','"
                        ,MCAE$X__8[i],"','"
                        ,MCAE$X__9[i],"','"
                        ,MCAE$X__10[i],"','"
                        ,MCAE$X__11[i],"','"
                        ,MCAE$X__12[i],"','"
                        ,MCAE$X__13[i],"','"
                        ,MCAE$X__14[i],"','"
                        ,MCAE$X__15[i],"','"
                        ,MCAE$X__16[i],"','"
                        ,MCAE$X__17[i],"','"
                        ,MCAE$X__18[i],"','"
                        ,MCAE$X__19[i],"','"
                        ,MCAE$X__20[i],"','"
                        ,MCAE$X__21[i],"','"
                        ,MCAE$X__22[i],"','"
                        ,MCAE$X__23[i],"','"
                        ,MCAE$X__24[i],"','"
                        ,MCAE$X__25[i],"','"
                        ,MCAE$X__26[i],"','"
                        ,MCAE$X__27[i],"','"
                        ,MCAE$X__28[i],"','"
                        ,MCAE$X__29[i],"','"
                        ,MCAE$X__30[i],"','"
                        ,MCAE$X__31[i],"','"
                        ,MCAE$X__32[i],"','"
                        ,MCAE$X__33[i],"','"
                        ,MCAE$X__34[i],"','"
                        ,MCAE$X__35[i],"','"
                        ,MCAE$X__36[i],"','"
                        ,MCAE$X__37[i],"','"
                        ,MCAE$X__38[i],"','"
                        ,MCAE$X__39[i],"','"
                        ,MCAE$X__40[i],"','"
                        ,MCAE$X__41[i],"','"
                        ,MCAE$X__42[i],"','"
                        ,MCAE$X__43[i],"','"
                        ,MCAE$X__44[i],"','"
                        ,MCAE$X__45[i],"','"
                        ,MCAE$X__46[i],"','"
                        ,MCAE$X__47[i],"','"
                        ,MCAE$X__48[i],"','"
                        ,MCAE$X__49[i],"','"
                        ,MCAE$X__50[i],"','"
                        ,MCAE$X__51[i],"','"
                        ,MCAE$X__52[i],"','"
                        ,MCAE$X__53[i],"','"
                        ,MCAE$X__54[i],"','"
                        ,MCAE$X__55[i],"','"
                        ,MCAE$CREATE_DATE[i],"');",sep="")
      tryCatch({dbSendQuery(mydb, MCAE_SQL)},finally = print(paste("insert data successful","MCAE", "on record No",i,sep=" ")))
      # i <- i+1

      dbClearResult(dbListResults(mydb)[[1]])
  }

  MCAE.Report <- dbSendQuery(mydb,"select * from mcae")
  data1 <- fetch(MCAE.Report, n = -1) 
  
  Guest.Data <- data1 %>% select(PROCESS_MONTH,CCARD_NO,TITLE,FIRST_NAME,LAST_NAME,GUESTS) %>% filter(as.numeric(GUESTS) > 1) %>% mutate(CNT = GUESTS,REASON = "GUESTS More than 2")
  
  CardHolder.Main.Data <- data1 %>% group_by(PROCESS_MONTH,CCARD_NO,TITLE,FIRST_NAME,LAST_NAME)  %>% summarise(CNT = n()) %>% filter(substr(CCARD_NO,10,10) != "9") %>% mutate(GUESTS = 0,REASON = "Main Card More than 5")
  
  CardHolder.Extra.Data <- data1 %>% group_by(PROCESS_MONTH,CCARD_NO,TITLE,FIRST_NAME,LAST_NAME)  %>% summarise(CNT = n()) %>% filter(substr(CCARD_NO,10,10) == "9") %>% mutate(GUESTS = 0,REASON = "Extra Card More than 5")
  
  Guest.Data <- Guest.Data %>% select (PROCESS_MONTH,CCARD_NO,TITLE,FIRST_NAME,LAST_NAME,REASON,CNT) %>% filter(substr(PROCESS_MONTH,1,7) == paste(Year.Now,month.now,sep = "-"))
  
  CardHolder.Main.Data <- CardHolder.Main.Data %>% select (PROCESS_MONTH,CCARD_NO,TITLE,FIRST_NAME,LAST_NAME,REASON,CNT) %>% filter(CNT > 5 & substr(PROCESS_MONTH,1,7) == paste(Year.Now,month.now,sep = "-"))
  
  CardHolder.Extra.Data <- CardHolder.Extra.Data %>% select (PROCESS_MONTH,CCARD_NO,TITLE,FIRST_NAME,LAST_NAME,REASON,CNT) %>% filter(CNT > 5 & substr(PROCESS_MONTH,1,7) == paste(Year.Now,month.now,sep = "-"))
  
  All.Data <- bind_rows(Guest.Data,CardHolder.Main.Data,CardHolder.Extra.Data)
  Month.Para <- paste(month.name,Year.Now,sep = "-")
  create.workbook(Month.Para)
  # temp  <- rbind(c("table_title", rep('', ncol(All.Data)-1)), # title
  #       rep('', ncol(All.Data)), # blank spacer row
  #       names(All.Data), # column names
  #       unname(sapply(All.Data, as.character))) # data
  # write.xlsx(temp, "temp MCAE.xlsx", sheet Name="Sheet1",
  #            col.names=FALSE, row.names=FALSE, append=FALSE)
  # 
  # wb <- loadWorkbook("temp MCAE.xlsx")
  # sheets <- getSheets(wb)
  # autoSizeColumn(sheets[[1]], colIndex=1:ncol(All.Data))
  # help("Border")
  # saveWorkbook(wb,"temp MCAE.xlsx")
  # data2 <- rbind(data2,data1 %>% select(PROCESS_MONTH,CCARD_NO,TITLE,FIRST_NAME,LAST_NAME,GUESTS) %>% filter(as.numeric(GUESTS) > 1))