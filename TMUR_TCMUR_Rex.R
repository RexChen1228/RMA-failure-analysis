# TMUR TCMUR devide into three part: "FE"/BE/Refurbish
# Only Provide FE to GSM

#UPDATE 20232H  (202111-202210)
#20231H  (202205-202204)

rm(list=ls())
options(java.parameters = "-Xmx1024g", stringsAsFactors=F)
library(data.table);library(RODBC);library(XLConnect);library(reshape);library(shiny);library(rCharts); library(stringr);library(ISOweek);library(zoo)
options(shiny.maxRequestSize = -1, RCHART_LIB = 'HighCharts')

#update original QC raw data convert to RDS (DL RDS)
##import data with RDS
#yearly data


dt <- read_feather("C:/Users/rex4_chen/Desktop/feather/DataClean_raw.dt_d.feather")
ori_raw.dt <- data.table(dt)
# 1. 找出時間欄位（Date 或 POSIXct 類型）
time_cols <- names(ori_raw.dt)[sapply(ori_raw.dt, function(col) inherits(col, "POSIXct") | inherits(col, "Date"))]

# 2. 找出非時間欄位
non_time_cols <- setdiff(names(ori_raw.dt), time_cols)

# 3. 替換非時間欄位的 NA 為空字串
ori_raw.dt[, (non_time_cols) := lapply(.SD, function(col) ifelse(is.na(col), "", col)), .SDcols = non_time_cols]

# 4. 替換時間欄位的 NA 為默認日期，例如 "1970-01-01"
ori_raw.dt[, (time_cols) := lapply(.SD, function(col) ifelse(is.na(col), as.Date("1970-01-01"), col)), .SDcols = time_cols]

setnames(ori_raw.dt, "CUSTOMER_COUNTRY_CODE_DESC","COUNTRY")

##SERVER - select product line ## Update 20212H
# # 7.1
#raw.dt<- ori_raw.dt[PRODUCT_LINE %in% c("SV","SF","SC","SK"),]
# # 7.2
raw.dt<- ori_raw.dt[PRODUCT_LINE %in% "SB",] #20
# #7.4 76249
#raw.dt<- ori_raw.dt[PRODUCT_LINE %in% c("SV","SF"),] #32

# # Sugon
# raw.dt <- raw.dt[PRODUCT_CUSTOMER%like% '曙光|DAWNING|SUMA', ]
# table(raw.dt$PRODUCT_CUSTOMER)

##### Fullfill those LMD_PART_GROUP==NA, but Exception Code = ^E & ^X, by ORG_MODEL_PART_NO >> PART_GROUP

table(raw.dt$LMD_PART_GROUP%in% c(""," ",NA))
table(!raw.dt$EXCEPTION_CODE%in% c(""," ",NA,"E25","X25"))
raw.dt <- raw.dt[(LMD_PART_GROUP %in% c(""," ",NA)& !EXCEPTION_CODE %in% c("","E25","X25")), PART_MAP := "Y"]
raw.dt <- raw.dt[(LMD_PART_GROUP %in% c(""," ",NA) & EXCEPTION_CODE %like% "^X"), PART_MAP := "CB"]
raw.dt_Ecode <- raw.dt[PART_MAP%in% "Y",]
table(raw.dt_Ecode$PART_MAP)
table(raw.dt_Ecode$ORG_MODEL_PART_NO %like% "^90SB")
# #Export the SWAP card and VLOOKUP the LMD_PART_GROUP by yearly raw data in Excel
# # raw.dt_ESC <- data.table(raw.dt_Ecode[ORG_MODEL_PART_NO %like% "^90SC",ORG_MODEL_PART_DESC])
# #Export the SWAP PARTS and MAPPING the LMD_PART_GROUP by yearly raw data through ORG_MODEL_PART_NO,ORG_MODEL_PART_DESC in Excel
# raw.dt_Ecode <- subset(raw.dt_Ecode,select=c(ORG_MODEL_PART_NO,ORG_MODEL_PART_DESC))
# raw.dt_Ecode <- raw.dt_Ecode[!duplicated(paste(ORG_MODEL_PART_NO,ORG_MODEL_PART_DESC))]
# wb_p <- loadWorkbook("/data/Server_GTSD/Amber/TMUR_TCMUR/Part_Mapping.xlsx", create=T)
# createSheet(object = wb_p, name = "20212H" )
# writeWorksheet(object = wb_p, data = raw.dt_Ecode, sheet=1)
# saveWorkbook(wb_p)

table(raw.dt$LMD_PART_GROUP)
# For those SWAP Parts directly
#SB
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SB"), PART_GROUP_MAP := "MAIN BOARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SW"), PART_GROUP_MAP := "MAIN BOARD"]

#SVFCK
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SV"), PART_GROUP_MAP := "MAIN BOARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SF"), PART_GROUP_MAP := "MAIN BOARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90-MSV"), PART_GROUP_MAP := "MAIN BOARD"]

raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^03A"), PART_GROUP_MAP := "MEMORY"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^03G"), PART_GROUP_MAP := "MEMORY"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SKM"), PART_GROUP_MAP := "MEMORY"]

raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^01"), PART_GROUP_MAP := "CPU"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SKU"), PART_GROUP_MAP := "CPU"]

raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^0A"), PART_GROUP_MAP := "POWER"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SKP"), PART_GROUP_MAP := "POWER"]

raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^03B"), PART_GROUP_MAP := "SSD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SKU"), PART_GROUP_MAP := "SSD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^17G"), PART_GROUP_MAP := "HDD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^19"), PART_GROUP_MAP := "HDD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90-S0"), PART_GROUP_MAP := "HDD"]

raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^041"), PART_GROUP_MAP := "CARD READER"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^14"), PART_GROUP_MAP := "CABLE"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^176"), PART_GROUP_MAP := "ODD"]

raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90YV"), PART_GROUP_MAP := "INSOURCING CARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SKC"), PART_GROUP_MAP := "OUTSOURCING CARD"]

raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^ASMB"), PART_GROUP_MAP := "SMALL BOARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^BP"), PART_GROUP_MAP := "SMALL BOARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^CB"), PART_GROUP_MAP := "SMALL BOARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^FPB"), PART_GROUP_MAP := "SMALL BOARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^MCB"), PART_GROUP_MAP := "SMALL BOARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^PCIE"), PART_GROUP_MAP := "SMALL BOARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^RE"), PART_GROUP_MAP := "SMALL BOARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^TPM"), PART_GROUP_MAP := "SMALL BOARD"]

raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^IOM"), PART_GROUP_MAP := "CARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^MCI"), PART_GROUP_MAP := "CARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^PDB"), PART_GROUP_MAP := "CARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^PEB"), PART_GROUP_MAP := "CARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^PEI"), PART_GROUP_MAP := "CARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^PEM"), PART_GROUP_MAP := "CARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^PIKE"), PART_GROUP_MAP := "CARD"]
raw.dt <- raw.dt[(PART_MAP %in% "Y" & ORG_MODEL_PART_NO %like% "^90SC" & ORG_MODEL_PART_DESC %like% "^ESC8"), PART_GROUP_MAP := "CARD"]


table(raw.dt$PART_GROUP_MAP)
table(raw.dt$PART_GROUP_MAP%in% "" & raw.dt$PART_MAP %in% "Y")

table(raw.dt$LMD_PART_GROUP)
raw.dt <- raw.dt[PART_MAP %in% "Y", LMD_PART_GROUP:=PART_GROUP_MAP]
table(raw.dt$LMD_PART_GROUP)

# Identification                 
raw.dt <- raw.dt[!duplicated(paste(RMA_NO,ORG_SN,SEQ_NO)),]
raw.dt <- raw.dt[,RMA_ORG:=(paste(RMA_NO,ORG_SN))]
raw.dt <- raw.dt[,MODEL_PROBLEM_REPLACE:=(paste(CFR_MODEL,PROBLEM_CODE,REPLACE_PART_NO))]
raw.dt <- raw.dt[,RMA_SN:=(paste(RMA_NO,SERIAL_NO))]

#time interval:201902-202001
raw.dt <- raw.dt[,":="(Close_M = format(as.Date(CLOSED_DATE), "%Y/%m"))]
raw.dt <- raw.dt[,":="(Close_Y = format(as.Date(CLOSED_DATE), "%Y"))]
raw.dt <- raw.dt[, ":="(week = paste('w', substr(ISOweek(CLOSED_DATE), 3, 4),substr(ISOweek(CLOSED_DATE), 7, 8), sep = ""))]
close_m = c("2023/11","2023/12","2024/01","2024/02","2024/03","2024/04","2024/05","2024/06","2024/07","2024/08","2024/09","2024/10") #
#"2022/05","2022/06","2022/07","2022/08","2022/09","2022/10","2022/11","2022/12","2023/01","2023/02","2023/03","2023/04"
#"2021/11","2021/12","2022/01","2022/02","2022/03","2022/04","2022/05","2022/06","2022/07","2022/08","2022/09","2022/10"
#"2021/05","2021/06","2021/07","2021/08","2021/09","2021/10","2021/11","2021/12","2022/01","2022/02","2022/03","2022/04"
#"2022/11","2022/12","2023/01","2023/02","2023/03","2023/04","2023/05","2023/06","2023/07","2023/08","2023/09","2023/10"
org <- raw.dt
raw.dt <- raw.dt[Close_M %in% close_m, ]

#COUNTRY_NAME<-colnames(raw.dt)
#EU RMA Center list
wb <- XLConnect::loadWorkbook("D:/PQA/COCO/TCMUR/EU self-repair RMA.xlsx")
EU_RMA <- XLConnect::readWorksheet(wb, sheet = 1, colTypes = "character")


###SERVER TMUR & TCMUR Frontend and Backend for Quotation (Combination)
#Return = FE + BE


#FE+BE
#Return = FE + BE
# raw.dt <- raw.dt[(RMA_TYPE %in% c("RT1","RT2","RT3","RT4","RT10") & 
#                     CUSTOMER_TYPE %in% c("ASUS- Dealer","ASUS- Disti","ASUS- End-user","ASUS- Operator",
#                                          "ASUS- Reseller","ASUS- Retailer","[EMPTY]","")), FEBE:="FE"]
# 
# raw.dt <- raw.dt[(RMA_TYPE %in% c('RT5','RT7','RT8','RT9') & 
#                     CUSTOMER_TYPE %in% c("ASUS- ASP", "ASUS-Hub","ASUS-RC","ASUS-SBU",
#                                          "ASUS-Warehouse","ASUS- Warehouse","ASUS-Intra.depart.")),FEBE:="BE"]
# raw.dt <- raw.dt[,!c(which(colnames(raw.dt)=="FEBE")),with = F]

# The original FE Condition
# FE<-raw.dt[(RMA_TYPE %in% c("RT1","RT2","RT4") & CUSTOMER_TYPE %in% c("ASUS- Dealer", "ASUS- Disti", "ASUS- End-user" ,"ASUS- Operator", "ASUS- Reseller", "ASUS- SI" ,"[EMPTY]", "ASUS-Retailer", "ASUS- Retailer", ""))| (RMA_TYPE == "RT10" & CUSTOMER_TYPE %in% c("ASUS- Dealer", "ASUS- Disti", "ASUS- End-user" ,"ASUS- Operator", "ASUS- Reseller", "ASUS- SI" ,"[EMPTY]", "ASUS-Retailer", "ASUS- Retailer", "","ASUS- RC")),]
# FE<-FE[PART_NO %like% "^90",]

raw.dt <- raw.dt[RMA_TYPE %in% c("RT1","RT2","RT3","RT4","RT10"), FEBE:="FE"]
raw.dt <- raw.dt[RMA_TYPE %in% c('RT5','RT7','RT8','RT9'), FEBE:="BE"]
table(raw.dt$FEBE)
#Exclude Transfer

raw.dt <- raw.dt[(RMA_TYPE_REASON %in% c("REAS-13","REAS-14","REAS-15","REAS21")| 
                    ITEM_RESULT %like% "TR"|REASON_CODE %in% "X05"), FEBE:="TR"]
table(raw.dt$FEBE)

#Remove CID(include OOW [ ,OOW:=RECEIVE_DATE - WARRANTY+END_DTAE])
raw.dt <- raw.dt[(FINAL_CODE=="CID"|REASON_CODE %like% "^C" |VIRTUAL_REASON_CODE %like% "^C"
                  |CID == "Y"|CID_OOW == "Y" |OOW == "Y"|ITEM_RESULT %like% 'CID'|ITEM_RESULT %like% 'OOW'), Filter:="CID/OOW"]
raw.dt <- raw.dt[,!c(which(colnames(raw.dt)=="CID")),with = F]
raw.dt <- raw.dt[Filter=="CID/OOW",":="(CID = "Y")]
raw.dt <- raw.dt[CID %in% NA, CID:="N"]
table(raw.dt$CID)

table(ori_raw.dt$RMA_TYPE_REASON)

###Other Condition
#Exclude Refurbish (duty is sales)
raw.dt <- raw.dt[RMA_TYPE_REASON %in% c("REAS-9","REAS-26"), RMA_TYPE_REA:="Refurbish"]  
#Exclude Sales return (duty is sales)
raw.dt <- raw.dt[RMA_TYPE_REASON %in% c("REAS-11","REAS-18"), RMA_TYPE_REA:="Sales return"]
#DOA
raw.dt <- raw.dt[(REASON_CODE %in% c("X06","E25","X25")|EXCEPTION_CODE %in% c("E25","X25")),RMA_TYPE_REA:="DOA"]
#Part fail/Part return/ part misjudge/replace due to same part fail in final test
raw.dt <- raw.dt[(REASON_CODE %in% "R15" | ACTION_CODE %in% c("F22","F21","R15")),RMA_TYPE_REA:="Part fail"]

table(raw.dt$RMA_TYPE_REA)
### Focus on FE

FE <- raw.dt[FEBE=="FE", ]
FE <- FE[!(RMA_CENTER %in% EU_RMA$EU_RMA_CENTER),]
#FE[is.na(FE)] <- ""
#FE <- FE[!(SUBMIT_LOC == "OK"),]
FE <- FE[SUBMIT_LOC %in% c('BAD',"BAD-SYS","RTV","SALE","SALE1","SCRAP","SUPPORT-BU",""," ", NA)]
FE <- FE[!(PRE_SALE_FLAG == "Y"),]
FE <- FE[!(CID == "Y"),]
FE <- FE[!(RMA_TYPE_REA %in% c("Refurbish","Sales return","DOA","Part fail")),]
FE <- FE[COUNTRY %in% "Türkiye", COUNTRY:="TURKEY"] #confirm country name with coco (Türkiye should be TURKEY)

#ACC Condition??
# FE <- FE[!(RMA_CENTER %in% c("广州信维")),]
# FE <- FE[!(PRODUCT_CUSTOMER%like% ('天津曙光|DAWNING INFORMATION INDUSTRY CO.,LTD|
#                                    TSINGHUA TONGFANG CO.LTD|"TONGFANG HONGKONG LIMITED|
#                                    LANGCHAO ELECTRONIC INFORMATION|
#                                    NANJING PANZHONG COMPUTER TECHNOLOGY C0. LTD.,' ]


#Separate TMUR vs TCMUR(parts-changed return only)

FE_N <-subset(FE,!(FE$LMD_PART_GROUP ==''))

FE_TMUR_D<-FE[!duplicated (FE$RMA_SN), TMUR_DENOMINATOR := "V"]
FE_TMUR_D<-FE_TMUR_D[is.na(TMUR_DENOMINATOR), TMUR_DENOMINATOR := ""]
FE_TCMUR_D<-FE_N[!duplicated (FE_N$RMA_SN), TCMUR_DENOMINATOR := "V"]
FE_TCMUR_D<-FE_TCMUR_D[is.na(TCMUR_DENOMINATOR), TCMUR_DENOMINATOR := ""]

TMUR_D <- FE_TMUR_D[TMUR_DENOMINATOR == "V",]
TCMUR_D <- FE_TCMUR_D[TCMUR_DENOMINATOR == "V",]

#1.by month & country
# dcast.data.table exports two columns: Close_M + COUNTRY; raw is by LMD_PART_GROUP
TMUR_TCMUR_N_M <- dcast.data.table(FE_N, Close_M + COUNTRY ~ LMD_PART_GROUP , length,value.var="RMA_CENTER") #country
TMUR_D_M<- dcast.data.table(TMUR_D, Close_M + COUNTRY ~. ,length,value.var="RMA_CENTER")#country
TCMUR_D_M<- dcast.data.table(TCMUR_D, Close_M + COUNTRY ~. ,length,value.var="RMA_CENTER")#country
setnames(TMUR_D_M,colnames(TMUR_D_M),c("Close_M","COUNTRY","Total_Qty"))
setnames(TCMUR_D_M,colnames(TCMUR_D_M),c("Close_M","COUNTRY","Total_Qty"))

TMUR_output_M <- merge(TMUR_TCMUR_N_M,TMUR_D_M, all=TRUE)
TMUR_output_M[is.na(TMUR_output_M)] <- 0
if(sum(is.na(TMUR_output_M$Total_Qty)) != 0){
  output[is.na(TMUR_output_M$Total_Qty),Total_Qty] <- 0 }

TCMUR_output_M <- merge(TMUR_TCMUR_N_M,TCMUR_D_M, all=TRUE)
TCMUR_output_M[is.na(TCMUR_output_M)] <- 0
if(sum(is.na(TCMUR_output_M$Total_Qty)) != 0){
  output[is.na(TCMUR_output_M$Total_Qty),Total_Qty] <- 0 }

#2.by year & country
# dcast.data.table exports two columns: Close_Y + COUNTRY; raw is by LMD_PART_GROUP
TMUR_TCMUR_N_Y <- dcast.data.table(FE_N, COUNTRY ~ LMD_PART_GROUP , length,value.var="RMA_CENTER")#country
TMUR_D_Y<- dcast.data.table(TMUR_D, COUNTRY ~. ,length,value.var="RMA_CENTER")#country
TCMUR_D_Y<- dcast.data.table(TCMUR_D, COUNTRY ~. ,length,value.var="RMA_CENTER")#country
setnames(TMUR_D_Y,colnames(TMUR_D_Y),c("COUNTRY","Total_Qty"))
setnames(TCMUR_D_Y,colnames(TCMUR_D_Y),c("COUNTRY","Total_Qty"))

TMUR_output_Y <- merge(TMUR_TCMUR_N_Y,TMUR_D_Y, all=TRUE)
TMUR_output_Y[is.na(TMUR_output_Y)] <- 0
if(sum(is.na(TMUR_output_Y$Total_Qty)) != 0){
  output[is.na(TMUR_output_Y$Total_Qty),Total_Qty] <- 0 }

TCMUR_output_Y <- merge(TMUR_TCMUR_N_Y,TCMUR_D_Y, all=TRUE)
TCMUR_output_Y[is.na(TCMUR_output_Y)] <- 0
if(sum(is.na(TCMUR_output_Y$Total_Qty)) != 0){
  output[is.na(TCMUR_output_Y$Total_Qty),Total_Qty] <- 0 }

#3.by year & Region
# dcast.data.table exports two columns: Region; raw is by LMD_PART_GROUP needed update
wb_r <- XLConnect::loadWorkbook('D:/PQA/Raw_Data/Mapping/Mapping_Table/Mapping_Region_20241H.xlsx')
Map_Region <- XLConnect::readWorksheet(wb_r, sheet = 1, colTypes = "character")
setnames(Map_Region, "REGION", "REGION_MAP")

TMUR_D_R <- merge(TMUR_D, Map_Region, by = "COUNTRY", all.x = TRUE)#country
TCMUR_D_R <- merge(TCMUR_D, Map_Region, by = "COUNTRY", all.x = TRUE)#country
FE_N_R <- merge(FE_N, Map_Region, by = "COUNTRY", all.x = TRUE)#country

table(FE_N_R$REGION_MAP %in%"")
table(FE_N_R$REGION_MAP %in%"",FE_N_R$COUNTRY)#country

TMUR_TCMUR_N_R <- dcast.data.table(FE_N_R, REGION_MAP ~ LMD_PART_GROUP , length,value.var="RMA_CENTER")
TMUR_D_R_Y<- dcast.data.table(TMUR_D_R, REGION_MAP ~. ,length,value.var="RMA_CENTER")
TCMUR_D_R_Y<- dcast.data.table(TCMUR_D_R, REGION_MAP ~. ,length,value.var="RMA_CENTER")
setnames(TMUR_D_R_Y,colnames(TMUR_D_R_Y),c("REGION_MAP","Total_Qty"))
setnames(TCMUR_D_R_Y,colnames(TCMUR_D_R_Y),c("REGION_MAP","Total_Qty"))

TMUR_output_R <- merge(TMUR_TCMUR_N_R,TMUR_D_R_Y, all=TRUE)
TMUR_output_R[is.na(TMUR_output_R)] <- 0
if(sum(is.na(TMUR_output_R$Total_Qty)) != 0){
  output[is.na(TMUR_output_R$Total_Qty),Total_Qty] <- 0 }

TCMUR_output_R <- merge(TMUR_TCMUR_N_R,TCMUR_D_R_Y, all=TRUE)
TCMUR_output_R[is.na(TCMUR_output_R)] <- 0
if(sum(is.na(TCMUR_output_R$Total_Qty)) != 0){
  output[is.na(TCMUR_output_R$Total_Qty),Total_Qty] <- 0 }

setnames(TMUR_output_R, "REGION_MAP", "COUNTRY")
setnames(TCMUR_output_R, "REGION_MAP", "COUNTRY")
length(TMUR_output_R)
#4. WW
#TMUR_WW_Matrix <- subset(TMUR_output_R,select=(c(2:35))) [follow TMUR_output_R n of variables SVCKF:33; SB:20;SVF:32]
TMUR_WW_Matrix <- subset(TMUR_output_R,select=(c(2:15)))
TMUR_WW_Matrix <- sapply(TMUR_WW_Matrix,as.numeric)
TMUR_WW_Matrix <- data.table(t(apply(TMUR_WW_Matrix, 2, sum)))
TMUR_output_WW <- TMUR_WW_Matrix[, COUNTRY :='WW'] 

TCMUR_WW_Matrix <- subset(TCMUR_output_R,select=(c(2:15)))
TCMUR_WW_Matrix <- sapply(TCMUR_WW_Matrix,as.numeric)
TCMUR_WW_Matrix <- data.table(t(apply(TCMUR_WW_Matrix, 2, sum)))
TCMUR_output_WW <- TCMUR_WW_Matrix[, COUNTRY :='WW']

TMUR_output_Y <- TMUR_output_Y[, Close_M :='2023/11-2024/10'] #2021/05-2022/04
TCMUR_output_Y <- TCMUR_output_Y[, Close_M :='2023/11-2024/10']

TMUR_output_R <- TMUR_output_R[, Close_M :='2023/11-2024/10']
TCMUR_output_R <- TCMUR_output_R[, Close_M :='2023/11-2024/10']

TMUR_output_WW <- TMUR_output_WW[, Close_M :='2023/11-2024/10']
TCMUR_output_WW <- TCMUR_output_WW[, Close_M :='2023/11-2024/10']

TMUR_list = list(TMUR_output_M,TMUR_output_Y,TMUR_output_R,TMUR_output_WW)
TMUR_output <- rbindlist(TMUR_list, use.names=TRUE, fill=TRUE)
TCMUR_list = list(TCMUR_output_M,TCMUR_output_Y,TCMUR_output_R,TCMUR_output_WW)
TCMUR_output <- rbindlist(TCMUR_list, use.names=TRUE, fill=TRUE)

#Creat Matrix dimension output
#TMUR_Matrix <- subset(TMUR_output,select=(c(3:39))) exclude: 1 date; 2 Country; the last one Total Q
TMUR_Matrix <- subset(TMUR_output,select=(c(3:15)))
TMUR_Matrix <- sapply(TMUR_Matrix,as.numeric)
TMUR_TQty <- as.numeric(TMUR_output$Total_Qty)

TCMUR_Matrix <- subset(TCMUR_output,select=(c(3:15)))
TCMUR_Matrix <- sapply(TCMUR_Matrix,as.numeric)
TCMUR_TQty <- as.numeric(TCMUR_output$Total_Qty)

TMUR_i <- TMUR_Matrix/TMUR_TQty
TCMUR_i <- TCMUR_Matrix/TCMUR_TQty

TMUR_g <- rowSums(TMUR_Matrix)/TMUR_TQty
TCMUR_g <- rowSums(TCMUR_Matrix)/TCMUR_TQty

TMUR_Percentage <- cbind(TMUR_i,TMUR_g)
TCMUR_Percentage <- cbind(TCMUR_i,TCMUR_g)

TMUR_Percentage <- t(apply(TMUR_Percentage*100, MARGIN=1,FUN=sprintf, fmt="%.2f%%")) 
TCMUR_Percentage <- t(apply(TCMUR_Percentage*100, MARGIN=1,FUN=sprintf, fmt="%.2f%%")) 

TMUR_Percentage <- data.table(TMUR_Percentage)
TCMUR_Percentage <- data.table(TCMUR_Percentage)

TMUR <- cbind(TMUR_output,TMUR_Percentage)
TCMUR <- cbind(TCMUR_output,TCMUR_Percentage)

WB_TMUR<-XLConnect::loadWorkbook('C:/Users/rex4_chen/Desktop/202311-202410_FE_TMUR_SB.xlsx',create=TRUE) ##create excel
createSheet(WB_TMUR, name = "TMUR DATA") ##create sheet
writeWorksheet(WB_TMUR, TMUR, sheet = "TMUR DATA")  
XLConnect::saveWorkbook(WB_TMUR)

WB_TCMUR<-XLConnect::loadWorkbook('C:/Users/rex4_chen/Desktop/202311-202410_FE_TCMUR_SB.xlsx',create=TRUE) ##create excel
createSheet(WB_TCMUR, name = "TCMUR DATA") ##create sheet
writeWorksheet(WB_TCMUR, TCMUR, sheet = "TCMUR DATA")  
XLConnect::saveWorkbook(WB_TCMUR)
# 
# ### Focus on Refurbish 
# 
# REF <-raw.dt[RMA_TYPE %in% c("RT4","RT7"),]
# #Check RT7 Containing
# table(REF$RMA_TYPE_REASON)
# table(REF$RMA_TYPE_REA)
# table(REF$RMA_TYPE_REASON %in% "REAS-6",REF$COUNTRY)
# REF <-REF[PART_NO %like% "^90",]
# REF <-REF[RMA_TYPE_REASON %in% c("REAS-9","REAS-26")|(COUNTRY =="JAPAN" & RMA_TYPE_REASON =="REAS-6"),]
# 
# REF <- REF[!(SUBMIT_LOC == "OK"),]
# REF <- REF[!(PRE_SALE_FLAG == "Y"),]
# REF <- REF[!(FEBE=="TR"), ]
# REF <- REF[!(RMA_TYPE_REA %in% "Part fail"),]
# REF <- REF[!(CID == "Y"),]
# 
# REF_N <- subset(REF,!(REF$LMD_PART_GROUP ==''))
# 
# REF_TMUR_D <- REF[!duplicated (REF$RMA_SN), TMUR_DENOMINATOR := "V"]
# REF_TMUR_D <-REF_TMUR_D[is.na(TMUR_DENOMINATOR), TMUR_DENOMINATOR := ""]
# REF_TCMUR_D <- REF_N[!duplicated (REF_N$RMA_SN), TCMUR_DENOMINATOR := "V"]
# REF_TCMUR_D <- REF_TCMUR_D[is.na(TCMUR_DENOMINATOR), TCMUR_DENOMINATOR := ""]
# 
# TMUR_D <- REF_TMUR_D[TMUR_DENOMINATOR == "V",]
# TCMUR_D <- REF_TCMUR_D[TCMUR_DENOMINATOR == "V",]
# 
# TMUR_TCMUR_N <- dcast.data.table(REF_N, Close_M + COUNTRY ~ LMD_PART_GROUP , length,value.var="RMA_CENTER")
# TMUR_D<- dcast.data.table(TMUR_D, Close_M + COUNTRY ~. ,length,value.var="RMA_CENTER")
# TCMUR_D<- dcast.data.table(TCMUR_D, Close_M + COUNTRY ~. ,length,value.var="RMA_CENTER")
# setnames(TMUR_D,colnames(TMUR_D),c("Close_M","COUNTRY","Total_Qty"))
# setnames(TCMUR_D,colnames(TCMUR_D),c("Close_M","COUNTRY","Total_Qty"))
# 
# REF_TMUR_output <- merge(TMUR_TCMUR_N,TMUR_D, all=TRUE)
# REF_TMUR_output[is.na(REF_TMUR_output)] <- 0
# if(sum(is.na(REF_TMUR_output$Total_Qty)) != 0){
#   output[is.na(REF_TMUR_output$Total_Qty),Total_Qty] <- 0 }
# 
# REF_TCMUR_output <- merge(TMUR_TCMUR_N,TCMUR_D, all=TRUE)
# REF_TCMUR_output[is.na(REF_TCMUR_output)] <- 0
# if(sum(is.na(REF_TCMUR_output$Total_Qty)) != 0){
#   output[is.na(REF_TCMUR_output$Total_Qty),Total_Qty] <- 0 }
# 
# #Creat Matrix dimension output
# #REF_TMUR_Matrix <- subset(REF_TMUR_output,select=(c(3:7)))[follow REF_TMUR_output n of variables SVCKF:7;SVF:7; SBW:11]
# REF_TMUR_Matrix <- subset(REF_TMUR_output,select=(c(3:11)))
# REF_TMUR_Matrix <- sapply(REF_TMUR_Matrix,as.numeric)
# REF_TMUR_TQty <- as.numeric(REF_TMUR_output$Total_Qty)
# 
# REF_TCMUR_Matrix <- subset(REF_TCMUR_output,select=(c(3:11)))
# REF_TCMUR_Matrix <- sapply(REF_TCMUR_Matrix,as.numeric)
# REF_TCMUR_TQty <- as.numeric(REF_TCMUR_output$Total_Qty)
# 
# REF_TMUR_i <- REF_TMUR_Matrix/REF_TMUR_TQty
# REF_TCMUR_i <- REF_TCMUR_Matrix/REF_TCMUR_TQty
# 
# REF_TMUR_g <- rowSums(REF_TMUR_Matrix)/REF_TMUR_TQty
# REF_TCMUR_g <- rowSums(REF_TCMUR_Matrix)/REF_TCMUR_TQty
# 
# REF_TMUR_Percentage <- cbind(REF_TMUR_i,REF_TMUR_g)
# REF_TCMUR_Percentage <- cbind(REF_TCMUR_i,REF_TCMUR_g)
# 
# REF_TMUR_Percentage <- t(apply(REF_TMUR_Percentage*100, MARGIN=1,FUN=sprintf, fmt="%.2f%%")) 
# REF_TCMUR_Percentage <- t(apply(REF_TCMUR_Percentage*100, MARGIN=1,FUN=sprintf, fmt="%.2f%%")) 
# 
# REF_TMUR_Percentage <- data.table(REF_TMUR_Percentage)
# REF_TCMUR_Percentage <- data.table(REF_TCMUR_Percentage)
# 
# REF_TMUR <- cbind(REF_TMUR_output,REF_TMUR_Percentage)
# REF_TCMUR <- cbind(REF_TCMUR_output,REF_TCMUR_Percentage)
# 
# WB_TMUR_REF<-loadWorkbook('/data/Server_GTSD/Amber/TMUR_TCMUR/201805-201904-TMUR-REF_SBW.xlsx',create=TRUE) ##create excel
# createSheet(WB_TMUR_REF, name = "TMUR_Refurbish") ##create sheet
# writeWorksheet(WB_TMUR_REF, REF_TMUR, sheet = "TMUR_Refurbish")  
# saveWorkbook(WB_TMUR_REF)
# 
# WB_TCMUR_REF<-loadWorkbook('/data/Server_GTSD/Amber/TMUR_TCMUR/201805-201904-TCMUR-REF_SBW.xlsx',create=TRUE) ##create excel
# createSheet(WB_TCMUR_REF, name = "TCMUR_Refurbish") ##create sheet
# writeWorksheet(WB_TCMUR_REF, REF_TCMUR, sheet = "TCMUR_Refurbish")  
# saveWorkbook(WB_TCMUR_REF)
# 

