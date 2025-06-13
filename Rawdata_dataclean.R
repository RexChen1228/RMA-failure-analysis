#L0-L2 去除workstation以及RT4 BU剔除ESB FEBE選擇FE 
QCRawData_All <- read_feather("C:/Users/rex4_chen/Desktop/feather/QCRawData_ALL.feather")
ori_raw.dt <- data.table(QCRawData_All)
raw.dt_d <- ori_raw.dt[!duplicated(ori_raw.dt),]
#合併ORG_SN以及RMA_NO---------------------------------------------------------------------------------
#處理SN2 
## Identification

raw.dt_d[,SN_RN:=(paste(ORG_SN,RMA_NO))]
raw.dt_d[,SN2:=substr(ORG_SN,1,2)]

mapping <- read_excel("D:/PQA/Raw_Data/Mapping/Mapping_Table_SN2.xlsx")#新的一年要新增
raw.dt_d[ ,SN2_YM:= mapping$SN2_DESC[match(SN2,mapping$SN2)]]
raw.dt_d[ ,SN2_Y := substr(raw.dt_d$SN2_YM,1,4) ]

raw.dt_d[,Summary_Code:= trimws(paste(ACTION_CODE,EXCEPTION_CODE))]
raw.dt_d[Summary_Code %chin% c(" ", ""), 
         Summary_Code := ifelse(Summary_Code == " ", REASON_CODE, RMA_TYPE_REASON)]

raw.dt_d[,Summary:=trimws(paste(ACTION_DESC,EXCEPTION_DESC))]
raw.dt_d[ Summary %chin% c(" ",""),
          Summary := ifelse(Summary == " ", REASON, RMA_TYPE_REASON_DESC)]

##PROBLEM
raw.dt_d[, Problem_Combine := PROBLEM]
raw.dt_d[is.na(Problem_Combine) | Problem_Combine == "" | Problem_Combine == " ", Problem_Combine := EXCEPTION_PROBLEM]
raw.dt_d[is.na(Problem_Combine) | Problem_Combine == "" | Problem_Combine == " ", Problem_Combine := `CUSTOMER_PROBLEM/CUSTOMER_MEMO`]
raw.dt_d[is.na(Problem_Combine) | Problem_Combine == "" | Problem_Combine == " ", Problem_Combine := RMA_TYPE_REASON_DESC]

table(raw.dt_d$Problem_Combine %in% "")

#Problem_Map!!

#Problem_Replace
raw.dt_d[,Problem_Replace_Code:=(paste(PROBLEM_CODE,REPLACE_PART_NO))]
raw.dt_d[,Problem_Replace:=(paste(Problem_Combine,REPLACE_PART_DESC))]

table(raw.dt_d$Problem_Replace %in% "")
#------------------------------------------------------------------------------------------------------

#workstation list----------------------------------------------------------------------------------------
WS_Mapping <- read_excel("D:/PQA/Raw_Data/Mapping/Workstation_list.xlsx")

setDT(WS_Mapping)

raw.dt_d[, WS := fifelse(
  `MODEL(PRODUCT_DESC)` %chin% WS_Mapping$Model | 
    PRODUCT_MODEL %chin% WS_Mapping$Model | 
    `ORG_MODEL(PRODUCT_DESC)` %chin% WS_Mapping$Model, 
  "Y", "N")]

raw.dt_d <- raw.dt_d[!(grepl("E500|E800|E900|^EG", `MODEL(PRODUCT_DESC)`))]


#先篩選條件
raw.dt_d <- raw.dt_d[WS == "N",]
raw.dt_d <- raw.dt_d[raw.dt_d$RMA_TYPE != "RT4", ]
raw.dt_d <- raw.dt_d[BU_TYPE != "ESB",]

#Filter
#CID(include OOW [ ,OOW:=RECEIVE_DATE - WARRANTY_END_DTAE])
raw.dt_d[(grepl("^C", REASON_CODE) | CID == "Y" | CID_OOW == "Y" | OOW == "Y"), Filter := "CID/OOW"]
raw.dt_d[grepl("^N", ACTION_CODE) | grepl("^N", REASON_CODE), Filter := "NTF"]

#VIRTUAL_REASON_CODE/ITEM_RESULT concern Pirority
raw.dt_d[, Filter := fcase(
  VIRTUAL_REASON_CODE %like% "^C" | ITEM_RESULT %like% "CID" | ITEM_RESULT %like% "OOW", "CID/OOW",
  VIRTUAL_REASON_CODE %like% "^N" | ITEM_RESULT %like% "^P", "NTF",
  is.na(Filter), "VID/Others"
)]


#FE+BE
#Return = FE + BE

#Mapping RMA Type
raw.dt_R <- raw.dt_d[RMA_TYPE %like% "^RT", RMA_TYPE_MAP:=RMA_TYPE]
raw.dt_R[RMA_TYPE %in% c("MR","O","R","U","Z"), RMA_TYPE_MAP:="RT1"]
raw.dt_R[RMA_TYPE %in% c("AR","D","X"), RMA_TYPE_MAP:="RT2"]
raw.dt_R[RMA_TYPE %in% c("SR","Y"), RMA_TYPE_MAP:="RT3"]
raw.dt_R[RMA_TYPE %in% c("P","PF"), RMA_TYPE_MAP:="RT4"]
raw.dt_R[RMA_TYPE %in% c("K","PR","Q","S","W"), RMA_TYPE_MAP:="RT5"]
raw.dt_R[RMA_TYPE %in% c("MT","SO"), RMA_TYPE_MAP:="RT9"]

table(raw.dt_R$RMA_TYPE_MAP)

table(raw.dt_R$RMA_TYPE_MAP =="RT9", raw.dt_R$RMA_TYPE_REASON)
table(raw.dt_R$RMA_TYPE_MAP =="RT9", raw.dt_R$RMA_TYPE_REASON_DESC)

#FE+BE
#Return = FE + BE
raw.dt_E <- raw.dt_R[RMA_TYPE_MAP %in% c("RT1","RT2","RT3","RT4","RT10","RT11"), FEBE:="FE"]
raw.dt_E[RMA_TYPE_MAP %in% c('RT5','RT7','RT8','RT9'),FEBE:="BE"]
#Exclude Transfer
raw.dt_E[(RMA_TYPE_REASON %in% c("REAS-13","REAS-14","REAS-15","REAS-21","REAS-32")|
            ITEM_RESULT %like% "TR"|REASON_CODE %in% "X05"), FEBE:="TR"]

table(raw.dt_E$FEBE)

#Create Variables
#Date
raw.dt_D <- raw.dt_E[, Close_M := format(as.Date(CLOSED_DATE), "%m")]
raw.dt_D[, Close_Y := format(as.Date(CLOSED_DATE), "%Y")]
raw.dt_D[, Close_YM := format(as.Date(CLOSED_DATE), "%Y/%m")]
raw.dt_D[Close_M %in% c("01","02","03"), Close_Q := "Q1"]
raw.dt_D[Close_M %in% c("04","05","06"), Close_Q := "Q2"]
raw.dt_D[Close_M %in% c("07","08","09"), Close_Q := "Q3"]
raw.dt_D[Close_M %in% c("10","11","12"), Close_Q := "Q4"]
raw.dt_D[, Close_YQ := paste(Close_Y,Close_Q,sep = "-")]

table(raw.dt_D$Close_M)
table(raw.dt_D$Close_Q)

#Model Name
table(raw.dt_D$ORG_MODEL_PART_DESC == "")
raw.dt_d <- raw.dt_D[ORG_MODEL_PART_DESC == "", ORG_MODEL_PART_DESC := PROJECT_NAME]
table(raw.dt_d$ORG_MODEL_PART_DESC == "")

#Mapping
wb_map <- XLConnect::loadWorkbook("D:/PQA/Raw_Data/Mapping/Mapping_Table/Mapping_Table_20250611.xlsx") #!!!update mapping table 手動更新
gc();xlcFreeMemory()
Map_Model <- data.table(readWorksheet(wb_map, sheet = "Mapping Table", colTypes = "character"))
Map_SYS_MB <- data.table(readWorksheet(wb_map, sheet = "SYS_MB", colTypes = "character"))
Map_SN <- data.table(readWorksheet(wb_map, sheet = "SN", colTypes = "character"))
MKT_Name <- data.table(readWorksheet(wb_map, sheet = "MKT_Name", colTypes = "character"))

#Merge                                                                                                                                                                     
raw.dt_M <- merge(raw.dt_d, Map_Model, by = "ORG_MODEL_PART_DESC", all.x = TRUE, allow.cartesian=TRUE )
raw.dt_M <- merge(raw.dt_M, Map_SN, by = "ORG_SN", all.x = TRUE )
raw.dt_M[MODEL_NAME.x %in% NA , MODEL_NAME.x := MODEL_NAME.y]
raw.dt_M[SKU.x %in% NA , SKU.x := SKU.y]
raw.dt_M[TYPE.x %in% NA , TYPE.x := TYPE.y]
setnames(raw.dt_M, "MODEL_NAME.x", "MODEL_NAME")
setnames(raw.dt_M, "SKU.x", "SKU")
setnames(raw.dt_M, "TYPE.x", "TYPE")
raw.dt_M <- raw.dt_M[, !c("MODEL_NAME.y", "SKU.y","TYPE.y")]
raw.dt_M <- merge(raw.dt_M, Map_SYS_MB, by = "MODEL_NAME", all.x = TRUE, allow.cartesian=TRUE )

raw.dt_Md <- raw.dt_M[!duplicated(raw.dt_M)]

table(raw.dt_Md$MODEL_NAME%in%NA)

on.exit(close(wb_map))#用来確保在你處理完 Excel 文件後關閉工作簿 wb_map，以釋放資源，防止内存泄漏或文件被鎖定。

#Export those MODEL_NAME is NA, and maintain the mapping table
raw.dt_NA <- data.table(raw.dt_Md[is.na(MODEL_NAME), .(ORG_MODEL_PART_DESC, ORG_MODEL_PART_NO)])
#raw.dt_NA <- data.table(raw.dt_Md[MODEL_NAME%in%NA,ORG_MODEL_PART_DESC,ORG_MODEL_PART_NO])
raw.dt_NAd <- raw.dt_NA[!duplicated(raw.dt_NA)] #!!!valid with na value with Mapping_NA_01

# 生成當前年月作為文件名稱和工作表名稱
year_month <- format(Sys.Date(), "%Y_%m")
file_path <- paste0("D:/PQA/Raw_Data/Mapping/Mapping/Mapping_NA_", year_month, ".xlsx")

# 加載或創建工作簿
wb_na <- XLConnect::loadWorkbook(file_path, create = TRUE)#!!!update mapping na table 手動更新
gc(); xlcFreeMemory()

# 創建新工作表並寫入數據
XLConnect::createSheet(object = wb_na, name = year_month)
writeWorksheet(object = wb_na, data = raw.dt_NAd, sheet = year_month)

# 保存工作簿
XLConnect::saveWorkbook(wb_na)
gc(); xlcFreeMemory()

#Rename MKT name
raw.dt_N <- raw.dt_Md
raw.dt_N[MODEL_NAME == "WS460T" , MODEL_NAME := "ESC300 G4"]
raw.dt_N[MODEL_NAME == "WS470T" , MODEL_NAME := "TS100-E10"]
raw.dt_N[MODEL_NAME == "WS660T" , MODEL_NAME := "ESC500 G4"]
raw.dt_N[MODEL_NAME == "WS660 SFF" , MODEL_NAME := "ESC500 G4 SFF"]
raw.dt_N[MODEL_NAME == "WS670 SFF" , MODEL_NAME := "ESC510 G4 SFF"]
raw.dt_N[MODEL_NAME == "WS690 SFF" , MODEL_NAME := "E500 G5 SFF"]
raw.dt_N[MODEL_NAME == "WS860T" , MODEL_NAME := "ESC700 G3"]
raw.dt_N[MODEL_NAME %in% c("WS880T","E700 G4") , MODEL_NAME := "ESC700 G4"]
raw.dt_N[MODEL_NAME == "WS690T" , MODEL_NAME := "E500 G5"]
raw.dt_N[MODEL_NAME == "WS980T" , MODEL_NAME := "E900 G4"]
raw.dt_N[MODEL_NAME == "WS720T" , MODEL_NAME := "PRO E500 G6"]
raw.dt_N[MODEL_NAME %in% "WS950T" , MODEL_NAME := "PRO E800 G4"] 
raw.dt_N[MODEL_NAME %in% "WS750T" , MODEL_NAME := "PRO E500 G7"]
#----------------------------------------------------------------------------------------------------------------------------------------------------
#多用一個欄位紀錄BE的資料FE_BE_TR-----------------------------------------------------------------------------------------------
#先以換料角度來看
#以LMD做為判別依據，當欄位為空直時用MUC填入
raw.dt_d <- raw.dt_N[LMD_PART_GROUP %in% c(""," ",NA), LMD_PART_GROUP:= MUC_MODULE]
#直接以換料的角度來看，當符合以下的換料就是L1/L2
raw.dt_d[LMD_PART_GROUP %in% c("BASEBOARD","BATTERY","BRACKET","CABLE","CARD","CARD READER","COIN CELL","CPU","DONGLE", "CHASSIS","MECHANICAL","PACKING","CASE",
                               "FAN","GPU CARD","HDD","HOLDER","INSOURCING CARD","KEYBOARD","KEYBOARD + MOUSE SET","MAIN BOARD","MAIN BOARD "," MAIN BOARD ",
                               "MECHANICAL","MEMORY","MEMORY-IC","MOUSE","MP BASE UNIT","MYLAR","ODD","ODD BEZEL","OUTSOURCING CARD","SPONGE",
                               "POWER","POWER CORD","RUBBER","SCREW","SHIELDING","SMALL BOARD","SSD","SUBMATERIAL","TRAY","THERMAL",
                               "VGA BOARD","90-DD","90-IG"), FEBE_N := "FE"]

#再來以維修等級來看，符合以下的summary code就是屬於L1/L2-------------------------------------------------------------------------------------------------------------------
#先篩選FE_BE_N為空的
raw.dt_d[FEBE_N %in% c(""," ",NA) & VIRTUAL_REASON_CODE %in% c("R01","R02","R03","R05","R06","R14","R16","R20","R22","R23","R26"), FEBE_N := "FE"]
#先篩選FE_BE_N為空的
raw.dt_d[FEBE_N %in% c(""," ",NA) & VIRTUAL_REASON_CODE %like% c("E^")|VIRTUAL_REASON_CODE %like% c("N^")|VIRTUAL_REASON_CODE %like% c("X^"), FEBE_N := "FE"]
#需要被分類再BE的但是因為上面的條件導致分類再FE，還要確認看看還有沒有
raw.dt_d[VIRTUAL_REASON_CODE %in% c("F08","F30","R07","F19","R25"), FEBE_N := "BE"]#R25為申請L1-L2但是為L3-L4的維修
table(raw.dt_d$FEBE_N)


#需要校正回歸的類別
raw.dt_d[LMD_PART_GROUP %in% c("CAPACITOR","CHIPSET(BGA)","CONN","DISCRETE","FLASH","IC","INDUCTOR","PROGRAMABLE","RESISTOR",
                               "TOOLS","CHIPSET","PLATE","COVER","PAD"), FEBE_N := "BE"]
#因為reason code & virtual reason code為空導致直接被填入，參考EXCEPTION_CODE的值
raw.dt_d[EXCEPTION_CODE %in% c("REAS-3"), FEBE_N := "BE"]
raw.dt_d[VIRTUAL_REASON_CODE %in% c("F08","F30","R07","F19","R25"), FEBE_N := "BE"]#R25為申請L1-L2但是為L3-L4的維修
raw.dt_d[VIRTUAL_REASON_CODE %in% c("F16","F03","F09","E18","E27","E34","F23","X18","E16","E14"), FEBE_N := "FE"]
raw.dt_d[FEBE_N %in%c(""," ",NA), FEBE_N := paste0(FEBE,"-NA")]

#加入Product_Line資訊
raw.dt_d[, FEBE_NP := paste0(FEBE_N,"-",PRODUCT_LINE)]

#區分資料為odm以及CHANNEL-------------------------------------------------------------------------------------------------------------------------------------------
ODM_data<-read_excel("D:/PQA/Raw_Data/CRR/ODM_OEM_List.xlsx", sheet = "List")
ODM_model<-as.vector(apply(ODM_data,1, function(x) gsub(" ","",x,fixed=T)))
raw.dt_d[MODEL_NAME %in% ODM_model,ODM_OEM:="ODM"]
raw.dt_d[!MODEL_NAME  %in% ODM_model & !is.na(MODEL_NAME),ODM_OEM:="channel"]

raw.dt_d[, CLOSED_DATE := as.Date(CLOSED_DATE, format="%Y/%m/%d")]

#Count
raw.dt_d[,SEQ_NO_N := as.numeric(SEQ_NO)]
raw.dt_d[,KBO_SEQ_N := as.numeric(KBO_SEQ)]
raw.dt_d[,SEQ_NO_C := sprintf("%02d",SEQ_NO_N)]
raw.dt_d[,KBO_SEQ_C := sprintf("%02d",KBO_SEQ_N)]
#Order
raw.dt_d[,Repair_Order := (paste0(SEQ_NO_C,KBO_SEQ_C))]

#Virtual Reason Code以Reason Code填滿--------------------------------------------------------------------------------------------------
#virtual Reason code會自動以Reason code填滿，Reason code的來源為exception code+action code，會有少數這個欄位為空的
raw.dt_d[VIRTUAL_REASON_CODE %in% c(""," ",NA), VIRTUAL_REASON_CODE:=REASON_CODE] 

#--------------------------------------------------------------------------------------------------------------------------------------

# #排除back end資料
# #資料先切為SB/SYS，只要SB都算為L3-L4的維修
# BE_1 <- raw_data[PRODUCT_LINE  %in% c("SB") & !VIRTUAL_REASON_CODE %in% "X05",] #撈取BE_data排除X05(後送維修)
# #由於這種切法，會把L3-L4的keyparts的維修做排除，因此要做兩次filter,把RT code為 BE的再篩選一次
# BE_2 <- raw_data[!PRODUCT_LINE  %in% c("SB") & !VIRTUAL_REASON_CODE %in% "X05" & RMA_TYPE %in% c('RT5','RT7','RT8','RT9'),] #篩選非SB的BE資料
# BE_Data <- rbind(BE_1,BE_2)
# #轉廠資料(TR)
# TR_Data <- raw_data[VIRTUAL_REASON_CODE %in% "X05"]
# TR_BE_Data <- rbind(BE_Data,TR_Data)
# setDT(TR_BE_Data)
#---------------------------------------------------------------------------------------------------------------------------------------

#!假如Virtual_Reason_code以及Reason_code同時為空可以替代的code為何???
#CRR計算為L0-L2，首先先排除L3-L4的資料，僅保留L0-L2--------------------------------------------------------------------------------------------------
#data <- data[RMA_TYPE%in%c("RT1","RT2","RT3","RT4","RT10","RT11") & !VIRTUAL_REASON_CODE %in% "X05" &!PRODUCT_LINE  %in% c("SB"), ]
#---------------------------------------------------------------------------------------------------------------------------------------

#Define service level------------------------------------------------------------------------------------------------------------------- 
#維修等級對應
mapping <- read_excel("D:/Reason_Levels_Mapping.xlsx", sheet =1)
#快速的區分法，沒有換料的按照service list做區分，service list要定期更新
raw.dt_d[ REPLACE_PART_NO %in% c(""," ",NA), REPAIR_LEVEL:= mapping$REPAIR_LEVEL[match(VIRTUAL_REASON_CODE,mapping$REASON_CODE)]]
#快速的區分法，有換料的直接被歸類成L2.1，如果REPALCED欄位不為空就直接歸類L2.1
raw.dt_d[!REPLACE_PART_NO %in% c(""," ",NA), REPAIR_LEVEL:="L2.1"]
#---------------------------------------------------------------------------------------------------------------------------------------------

# 針對 R25 & COUNTRY %in% c("CHINA","TAIWAN")下，若 MUC_MODULE為MB(原本為空白 已替換成MB) 須將level 重新定義成L2.1  避免依照上面之Virtual Reason code判斷成L2.2
# 因為此狀況實際維修是換MAIN BOARD , 只是是維修MB上面的零件後直接換上去, 並非按照一般流程: 換下壞的MB 直接換新的MB
#!202308待觀察Server是否有這個狀況(目前僅有一筆，也確實符合上面條件)

raw.dt_d[VIRTUAL_REASON_CODE=="R25" & CUSTOMER_COUNTRY_CODE_DESC %in% c("CHINA","TAIWAN") & is.na(LMD_PART_GROUP),LMD_PART_GROUP:="MAIN BOARD"]
raw.dt_d[VIRTUAL_REASON_CODE=="R25" & CUSTOMER_COUNTRY_CODE_DESC %in% c("CHINA","TAIWAN") & LMD_PART_GROUP=="MAIN BOARD",REPAIR_LEVEL:="L2.1"]
#---------------------------------------------------------------------------------------------------------------------------------------------

#!以下邏輯待確認,由於各產品線所定義的90料號不一樣
#data <- data[LMD_PART_GROUP%like%"^90-", REPAIR_LEVEL:="L1.2"]

#0922 update
#!需重新定義哪些類別的維修是否按照RT code
#L1.1(Accessory SWAP/Exchange)，為配件相關，server的配件為? MUC_MODULE與LMD Part Group(細項較多)有差別
raw.dt_d[LMD_PART_GROUP %in% c("ADAPTER","ACCESSORY","MOUSE","KEYBOARD + MOUSE SET","KEYBOARD"), REPAIR_LEVEL:="L1.1"] 
#----------------------------------------------------------------------------------------------------------------------------------------------------------
#前面已經給過Repair level的值跳過，以virtual_reason_code為分類基準
raw.dt_d[REPAIR_LEVEL%in% c(""," ",NA) & REPLACE_PART_NO %in% c(""," ",NA) & VIRTUAL_REASON_CODE %in% c(""," ",NA) & RMA_TYPE=="RT3",  REPAIR_LEVEL:="L0"] #credit的service level等於L0
raw.dt_d[REPAIR_LEVEL%in% c(""," ",NA) & REPLACE_PART_NO %in% c(""," ",NA) & VIRTUAL_REASON_CODE %in% c(""," ",NA) & RMA_TYPE=="RT4",  REPAIR_LEVEL:="L2.1"]
raw.dt_d[REPAIR_LEVEL%in% c(""," ",NA) & REPLACE_PART_NO %in% c(""," ",NA) & VIRTUAL_REASON_CODE %in% c(""," ",NA) & RMA_TYPE=="RT1",  REPAIR_LEVEL:="L2.1"]
raw.dt_d[REPAIR_LEVEL%in% c(""," ",NA) & REPLACE_PART_NO %in% c(""," ",NA) & VIRTUAL_REASON_CODE %in% c(""," ",NA) & RMA_TYPE=="RT2",  REPAIR_LEVEL:="L1.1"]#待確認
raw.dt_d[REPAIR_LEVEL%in% c(""," ",NA) & REPLACE_PART_NO %in% c(""," ",NA) & VIRTUAL_REASON_CODE %in% c(""," ",NA) & RMA_TYPE=="RT10", REPAIR_LEVEL:="L0"]
#-----------------------------------------------------------------------------------------------------------------------------------------------------------

# Define CID & OOW as CID/OOW-----------------------------------------------------------------------------------------------------------------
#重新分類一次CID/OOW
raw.dt_d[CID %in% "Y" | OOW %in% "Y" |VIRTUAL_REASON_CODE %like% "C^", REPAIR_LEVEL:="CID/OOW"]

#---------------------------------------------------------------------------------------------------------------------------------------------

# Define Level as L0 if REPAIR_LEVEL is NULL
raw.dt_d[REPAIR_LEVEL %in% c(""," ",NA), REPAIR_LEVEL:="L0"]

#0519 將L1.1 ADAPTER 暫時改成 L2.1 ADAPTER]
#後來發現這樣的寫法 可能遇到單存只有 L2.2 L1.1 Adapter的狀況時 會保留一筆多修為 L1.1 ADAPTER 
#建議修改方法可用 by RMA_ORG 每個一筆多修的狀態(若要調整的話 否則如果覺得L1.1 ADTATER 比 L2.2重要則不需要調整)
#去判斷若沒有任何L2.1 則L1.1 ADATPER不需要做修改 (也就是此時可能只有"CID/OOW", "L1.2", "L2.2", "L1.1", "L1.3", "L0" 之狀況)
#去判斷若有存在L2.1 則L1.1 ADATPER需要修改成L2.1  (也就是此時可能只有"CID/OOW", "L1.2","L2.1, "L2.2", "L1.1", "L1.3", "L0" 之狀況)
#先將歸類成ADPATER,svr沒有之後有遇到再調整
raw.dt_d[LMD_PART_GROUP %in% c("ADAPTER") & REPAIR_LEVEL=="L1.1",REPAIR_LEVEL:="L2.1"]


#--------------------------------------------------------------------------------------------------------
# Sorting Data，為了要使用duplciated函數，sorting過後如果有重複值會自動保留第一筆
level.seq.v <- c("CID/OOW", "L1.2", "L2.1", "L2.2", "L1.1", "L1.3", "L0")

#先按照LMD Parts Group按照便宜到貴的類別作排序 
LMD_PART_GROUP.seq.v<- c("BASEBOARD","MAIN BOARD","CPU","GPU CARD","OUTSOURCING CARD","MEMORY","SSD","HDD","MEMORY-IC",
                         "CHIPSET(BGA)","PROGRAMABLE", "CARD", "VGA BOARD", "POWER", "CHASSIS", "NODE", 
                         "MECHANICAL", "FAN", "SMALL BOARD", "IC", "CPU", "CARD READER", "CHIPSET","FLASH",
                         "CONN", "CABLE","PACKING","THERMAL", "DISCRETE", "POWER CORD","PLATE","COIN CELL","MYLAR",
                         "CAPACITOR", "RESISTOR","INDUCTOR", "HOLDER", "BRACKET","IO SHIELD", "90-DD","90-LA","90-SB","90-SF","90-SK","BATTERY",
                         "MP BASE UNIT","SHIELDING","SUBMATERIAL","KEYBOARD","TRAY","SCREW","TOOLS","ODD","ODD BEZEL","KEYBOARD + MOUSE SET",
                         "MOUSE","SPONGE")


# Sorting data by Level as the first order, and by Module as the second order.
raw.dt_d$REPAIR_LEVEL <- factor(raw.dt_d$REPAIR_LEVEL, levels=level.seq.v)

#data$MUC_MODULE   <- factor(data$MUC_MODULE,   levels=MUC_MODULE.seq.v)

raw.dt_d <- raw.dt_d[order(REPAIR_LEVEL, LMD_PART_GROUP), ]

#把L2.1照 LMD_PART_GROUP.seq.v的順序排(先用Match找出對應的index在m先用Match找出對應的MUC_MODULE.seq.v index,有了index後
#再用order找出最小的index的位置抓出來，即可依照想要的LMD_PART_GROUP.seq.v順序做排序
raw.dt_d[REPAIR_LEVEL %in% "L2.1"] <- raw.dt_d[REPAIR_LEVEL %in% "L2.1"][order(match(LMD_PART_GROUP,LMD_PART_GROUP.seq.v))]

#0519 將L2.1 ADAPTER 改回 L1.1 ADAPTER ，此時L1.1 ADAPTER 順序已在 L2.1 CABLE之前 這樣之後判別一筆多修的時候 會保留 L1.1 ADAPTER 
#data[LMD_PART_GROUP %in% c("ADAPTER") & REPAIR_LEVEL=="L2.1",REPAIR_LEVEL:="L1.1"]#因為adpater被歸類在配件，由於判斷的關係會被歸類在L21，要

#data <- data[ ,`REPAIR_LEVEL(with multi repair flag)`:=REPAIR_LEVEL]
#data <- data[duplicated(paste0(RMA_NO,SERIAL_NO)), `REPAIR_LEVEL(with multi repair flag)`:="Multi Repair Data"]
raw.dt_d[duplicated(SN_RN), Mix_Repair_LMD := "Mix_Repair" ]
raw.dt_d[Mix_Repair_LMD %in% NA, Mix_Repair_LMD := "Main_Repair" ]

raw.dt_d <- raw.dt_d[order(SN_RN)]
raw.dt_d <- raw.dt_d[order(-Repair_Order)]

raw.dt_d[REPAIR_MEMO %in% "ARN record, created by System", Mix_Repair := "Do_ECN/ARN_Only"]
raw.dt_d[duplicated(SN_RN), Mix_Repair := "Mix_Repair" ]
raw.dt_d[Mix_Repair %in% NA, Mix_Repair := "Main_Repair" ]

table(raw.dt_d$Mix_Repair)

#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#加入L6/L10 Level-------------------------------------------------------------------------------------------------------------------------------------
raw.dt_d[PRODUCT_LINE %in% c("S","SV","SF") & `PART_DESC(MODEL)`%like% c("WOCPU"), L6_L10 := "準系統"]
Server_Level <- read_excel("D:/PQA/Raw_Data/Mapping/Mapping_Table_SYS_Level.xlsx", sheet = "SYS_Level")
raw.dt_d[,L6_L10 := Server_Level$SYS_Level[match(ORG_MODEL_PART_NO,Server_Level$Part_Number)]]
raw.dt_d[ L6_L10 %in%NA,  L6_L10 := TYPE]
#輸出RMA DATA沒有MAPPING到90 model name
raw.dt_SYS_NA <- data.table(raw.dt_d[TYPE %in% c("System") & L6_L10 %in% c("System"), .(ORG_MODEL_PART_NO, TYPE, L6_L10)])
raw.dt_SYS_NAd <- raw.dt_SYS_NA[!duplicated(raw.dt_SYS_NA)]

# 生成當前年月作為文件名稱
year_month <- format(Sys.Date(), "%Y_%m")
file_path <- paste0("C:/Users/rex4_chen/Desktop/Sys_Level_NA-", year_month, ".xlsx")

# 檢查資料是否存在，如果有資料則寫入 Excel
if (nrow(raw.dt_SYS_NAd) > 0) {
  openxlsx::write.xlsx(raw.dt_SYS_NAd, file = file_path, 
                       asTable = FALSE, rowNames = FALSE, overwrite = TRUE, sheetName = year_month)
}


raw.dt_d[(REASON_CODE %in% c("X06","E25","X25")|ACTION_CODE %in% c("E25","X25")|ITEM_RESULT%like%"-DOA"),RMA_TYPE_REA:="DOA"]

#Part fail/Part return/ part misjudge/replace due to same part fail in final test
raw.dt_d[ACTION_CODE %in% c("F22","F21","R15"),RMA_TYPE_REA:="Part fail"]

table(raw.dt_d$RMA_TYPE_REA)

raw.dt_d[, EMS_Shipout_Repair := as.Date(RECEIVE_DATE) - as.Date(RECEIVE_SN_EMS_SHIPOUT_DATE)]
#判別DOA
SN_F16_all <- raw.dt_d[VIRTUAL_REASON_CODE %in% c("R16","F16"), unique(ORG_SN)]

RMA_F16_all <- raw.dt_d[VIRTUAL_REASON_CODE %in% c("R16","F16"), unique(RMA_NO)]

SN_F16_follows_DOA <- raw.dt_d[!RMA_NO %in% RMA_F16_all & CUSTOMER_DOA=="Y" & ORG_SN %in% SN_F16_all, unique(ORG_SN)]
#給定Flag
raw.dt_d[RMA_NO %in% RMA_F16_all & ORG_SN %in% SN_F16_follows_DOA, R16_F16_DOA:="Y"]

raw.dt_d[EMS_Shipout_Repair <= 90, DOA_90 := "Y"]
raw.dt_d[EMS_Shipout_Repair > 90, DOA_90 := "N"]
raw.dt_d[EMS_Shipout_Repair <= 120, DOA_120 := "Y"]
raw.dt_d[EMS_Shipout_Repair > 120, DOA_120 := "N"]
raw.dt_d[EMS_Shipout_Repair <= 150, DOA_150 := "Y"]
raw.dt_d[EMS_Shipout_Repair > 150, DOA_150 := "N"]
raw.dt_d[EMS_Shipout_Repair <= 180, DOA_180 := "Y"]
raw.dt_d[EMS_Shipout_Repair > 180, DOA_180 := "N"]
raw.dt_d[REGION %in% c("EAST EURP","WEST EURP"), REGION := "EMEA"]
#BU的DOA定義-----------------------------------------------------------------------------------------------------------------------------------------------------------
raw.dt_d[(REASON_CODE %in% c("X06","E25","X25")|ACTION_CODE %in% c("E25","X25")|ITEM_RESULT%like%"-DOA"), X25_E25_ITEM_RESULT_DOA:="Y"]
raw.dt_d[!REPLACE_PART_DESC %in% c(""," ",NA) & LMD_PART_GROUP %in% c("MAIN BOARD")& REGION %in% c("EMEA","NA") & EMS_Shipout_Repair <= 14, EMS_Shipout_14_DOA:="Y"]
raw.dt_d[CUSTOMER_DOA=="Y",X25_E25_ITEM_RESULT_DOA:="Y"]
raw.dt_d[X25_E25_ITEM_RESULT_DOA =="Y" | EMS_Shipout_14_DOA =="Y" , extensive_DOA := "Y"]
raw.dt_d[is.na(X25_E25_ITEM_RESULT_DOA), X25_E25_ITEM_RESULT_DOA:= "N"]
raw.dt_d[is.na(EMS_Shipout_14_DOA), EMS_Shipout_14_DOA:= "N"]
raw.dt_d[is.na(extensive_DOA), extensive_DOA:= "N"]
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
raw.dt_d[, PRODUCT_LINE2 := fifelse(PRODUCT_LINE == "SB", "Server MB",
                                    fifelse(PRODUCT_LINE == "SC", "Server Card",
                                            fifelse(PRODUCT_LINE %in% c("SF", "SV"), "SERVER SYSTEM",
                                                    fifelse(PRODUCT_LINE == "SK", "Server Keyparts",
                                                            fifelse(PRODUCT_LINE == "SW", "SW-WORKSTATION MB", "NA")))))]
#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
raw.dt_d[R16_F16_DOA == "Y" | X25_E25_ITEM_RESULT_DOA == "Y" | EMS_Shipout_14_DOA == "Y", DOA_Case := "Y"]
raw.dt_d[is.na(DOA_Case), DOA_Case:="N"]
#Add_Duty--------------------------------------------------------------------------------------------------------------------------------------------------------------- 
raw.dt_d[PROBLEM%in%c(""," ",NA), PROBLEM:=EXCEPTION_PROBLEM]
raw.dt_d[ ,Category:="OTHERS"]
raw.dt_d[ ,Duty    :="OTHERS"]
#-----------------------------------------------------------------------------------
#REPAIR_LEVEL = L0--------------------------------------------------------------------------------------------------
# 1. L0 DOA Credit w/ problem-------------------------------------------------------
#No Trouble Found and Customer decides not to repair
raw.dt_d[REPAIR_LEVEL%in%"L0" & VIRTUAL_REASON_CODE%like%"^X" & 
           !PROBLEM%like%"^Test ok|^Test OK|Test ok/Cannot duplicate|Test ok, No trouble found" &
           !PROBLEM%in%c(""," ",NA) & extensive_DOA == "Y" ,
         Category:="L0 DOA Credit w/ problem"]
raw.dt_d[Category=="L0 DOA Credit w/ problem", Duty:="EMS/SQE"]
#------------------------------------------------------------------------------------

# 2. L0 DOA test ok & NTF-------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L0" & VIRTUAL_REASON_CODE%like%"^X" &
           PROBLEM%like%"^Test ok|^Test OK|Test ok/Cannot duplicate|Test ok, No trouble found" &
           extensive_DOA == "Y",
         Category:="L0 DOA test ok & NTF"]

raw.dt_d[REPAIR_LEVEL%in%"L0" & VIRTUAL_REASON_CODE%like%"^N" & 
           extensive_DOA == "Y",
         Category:="L0 DOA test ok & NTF"]

raw.dt_d[Category=="L0 DOA test ok & NTF", Duty:="SPM"]

raw.dt_d[REPAIR_LEVEL%in%"L0" & VIRTUAL_REASON_CODE == "X25" &
           PROBLEM%like%"^Test ok|^Test OK|Test ok/Cannot duplicate|Test ok, No trouble found" &
           extensive_DOA == "Y" & Category == "L0 DOA test ok & NTF", Duty:="EQM/SQA"]
#-------------------------------------------------------------------------------------

# 3. L0 DOA其他-----------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L0" & VIRTUAL_REASON_CODE%like%"^R|^F" & 
           extensive_DOA == "Y",
         Category:="L0 DOA Others"]

raw.dt_d[REPAIR_LEVEL%in%"L0" & VIRTUAL_REASON_CODE%like%"^X" & 
           PROBLEM%in%c(""," ",NA) &
           extensive_DOA == "Y",
         Category:="L0 DOA Others"]
#-------------------------------------------------------------------------------------

# 4. L0 RMA Others (NTF, 客人不修.)----------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L0" & extensive_DOA == "N",
         Category:="L0 RMA Others (Customer decides to stop repairing)"]

raw.dt_d[Category=="L0 RMA Others (Customer decides to stop repairing)", Duty:="CSC"]
#-----------------------------------------------------------------------------------------

# 5. L0 RMA 因品質問題退費(Credit back, due to quality issues)-----------------------------
raw.dt_d[REPAIR_LEVEL%in%"L0" & VIRTUAL_REASON_CODE%in%"X23" & extensive_DOA =="N",
         Category:="L0 RMA Credit back, due to quality issues"]
raw.dt_d[Category=="L0 RMA Credit back, due to quality issues", Duty:="CSC"]
#------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------

# 6. L1.1 Accessory-配件不會是Power RD(keyboard/mouse)-------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L1.1", Category:="L1.1 Accessory"]
raw.dt_d[REPAIR_LEVEL%in%"L1.1", Duty:="RD/SQA"]
#-------------------------------------------------------------------------------------------

# 7. L1.2 SWAP------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L1.2", Category:="L1.2 SWAP"]
raw.dt_d[REPAIR_LEVEL%in%"L1.2", Duty:="SPM"]
#-------------------------------------------------------------------------------------------

# 8. L1.3 DOA Credit w/ problem-------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L1.3" & VIRTUAL_REASON_CODE%like%"^R" & 
           !PROBLEM%like%"^Test ok|^Test OK|Test ok/Cannot duplicate|Test ok, No trouble found" &
           !PROBLEM%in%c(""," ",NA) & extensive_DOA =="Y",
         Category:="L1.3 DOA Credit w/ problem"]
raw.dt_d[Category=="L1.3 DOA Credit w/ problem", Duty:="SW RD"]
#-------------------------------------------------------------------------------------------

# 9. L1.3 DOA 其他--------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L1.3" & VIRTUAL_REASON_CODE%like%"^R" & 
           PROBLEM%in%c(""," ",NA) & extensive_DOA =="Y",
         Category:="L1.3 DOA Others"]
raw.dt_d[Category=="L1.3 DOA Others", Duty:="CSC"]
#-------------------------------------------------------------------------------------------

# 10. L1.3 RMA------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L1.3" & VIRTUAL_REASON_CODE%like%"^R" &
           extensive_DOA =="N",
         Category:="L1.3 RMA"]
raw.dt_d[Category=="L1.3 RMA", Duty:="SW/CSC"]
#-------------------------------------------------------------------------------------------

# 11. L1.3 DOA test ok & NTF----------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L1.3" & VIRTUAL_REASON_CODE%like%"^R" & 
           PROBLEM%like%"^Test ok|^Test OK|Test ok/Cannot duplicate|Test ok, No trouble found" &
           extensive_DOA =="Y",
         Category:="L1.3 DOA test ok & NTF"]
raw.dt_d[Category=="L1.3 DOA test ok & NTF", Duty:="SPM"]
#-------------------------------------------------------------------------------------------

# 12. L2.1 CSC-------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & VIRTUAL_REASON_CODE%like%"^N",
         Category:="L2.1-NTF"]
raw.dt_d[Category=="L2.1-NTF", Duty:="CSC"]
#-------------------------------------------------------------------------------------------

# 13. L2.1 MB------------------------------------------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE == "MAIN BOARD",
         Category:="L2.1-MB"]

raw.dt_d[Category=="L0 RMA Others (NTF, Customer decides not to repair.)" & VIRTUAL_REASON_CODE=="R25", Category:="L2.1-MB"]

raw.dt_d[REPAIR_LEVEL=="L2.2" & VIRTUAL_REASON_CODE=="R14" & CUSTOMER_COUNTRY_CODE_DESC=="CHINA", Category:="L2.1-MB"]
raw.dt_d[REPAIR_LEVEL=="L2.2" & VIRTUAL_REASON_CODE=="R25" & CUSTOMER_COUNTRY_CODE_DESC=="CHINA", Category:="L2.1-MB"]

raw.dt_d[Category=="L2.1-MB", Duty:="EE RD"]
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#36.先將 Power IC   PL, PC, PU, PQ, PD, PCE----------------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           REPLACE_PART_LOC %like%"PL",
         Category:="L2.1-POWER-PCB"]
raw.dt_d[Category=="L2.1-POWER-PCB", Duty:="Power RD"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           REPLACE_PART_LOC %like%"PC",
         Category:="L2.1-POWER-PCB"]
raw.dt_d[Category=="L2.1-POWER-PCB", Duty:="Power RD"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           REPLACE_PART_LOC %like%"PU",
         Category:="L2.1-POWER-PCB"]
raw.dt_d[Category=="L2.1-POWER-PCB", Duty:="Power RD"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           REPLACE_PART_LOC %like%"PQ",
         Category:="L2.1-POWER-PCB"]
raw.dt_d[Category=="L2.1-POWER-PCB", Duty:="Power RD"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           REPLACE_PART_LOC %like%"PD",
         Category:="L2.1-POWER-PCB"]
raw.dt_d[Category=="L2.1-POWER-PCB", Duty:="Power RD"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           REPLACE_PART_LOC %like%"PR",
         Category:="L2.1-POWER-PCB"]
raw.dt_d[Category=="L2.1-POWER-PCB", Duty:="Power RD"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           REPLACE_PART_LOC %like%"PCE" & !MUC_MODULE == "CONN",
         Category:="L2.1-POWER-PCB"]
raw.dt_d[Category=="L2.1-POWER-PCB", Duty:="Power RD"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           REPLACE_PART_LOC %like%"PQ",
         Category:="L2.1-POWER-PCB"]
raw.dt_d[Category=="L2.1-POWER-PCB", Duty:="Power RD"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           REPLACE_PART_LOC %like%"PU",
         Category:="L2.1-POWER-PCB"]
raw.dt_d[Category=="L2.1-POWER-PCB", Duty:="Power RD"]
#--------------------------------------------------------------------------------------------------------------------------------------------------------
# 14. L2.1 ME------------------------------------------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("CHASSIS","SCREW","ODD BEZEL","MECHANICAL","MECHANICAL-A","MECHANICAL-B","MECHANICAL-C","MECHANICAL-D","EXPENSE","PLATE","TRAY","HOLDER"),
         Category:="L2.1-ME"]

raw.dt_d[Category=="L2.1-ME", Duty:="ME RD"]
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# 13. L2.1 PANEL--------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL %in% "L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE %in% c("PANEL","LCD PANEL","LCD MODULE BKT","OLED PANEL","OLED MODULE"),
         Category:="L2.1-PANEL"]
raw.dt_d[Category=="L2.1-PANEL", Duty:="RD/SQA"]
#-----------------------------------------------------------------------------------------------------------------------------------------

# 14. L2.1 KEYBOARD-----------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE =="KEYBOARD",
         Category:="L2.1-KEYBOARD"]
raw.dt_d[Category=="L2.1-KEYBOARD", Duty:="RD/SQA"]
#-----------------------------------------------------------------------------------------------------------------------------------------

# 15. L2.1 HDD----------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE=="HDD",
         Category:="L2.1-HDD"]

raw.dt_d[Category=="L2.1-HDD", Duty:="RD/SQA"]
#-----------------------------------------------------------------------------------------------------------------------------------------
# 15. L2.1 SSD----------------------------------------------------------------------------------------------------------------------------
raw.dt_d[Category=="L2.1-HDD" & MUC_MODULE=="SSD",
         Category:="L2.1-SSD"]
raw.dt_d[Category=="L2.1-SSD", Duty:="RD/SQA"]
#-----------------------------------------------------------------------------------------------------------------------------------------

# 16. L2.1 BATTERY------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in% c("BATTERY","COIN CELL"),
         Category:="L2.1-BATTERY"]
raw.dt_d[Category=="L2.1-BATTERY", Duty:="RD/SQA"]
#-----------------------------------------------------------------------------------------------------------------------------------------

# 17. Power RD----------------------------------------------------------------------------------------------------------------------------
#分為PSU以及Power IC
#1. PSU
raw.dt_d[Category=="L2.1-POWER", Duty:="Power RD"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE=="POWER",
         Category:="L2.1-POWER"]
raw.dt_d[Category=="L2.1-POWER", Duty:="Power RD"]

#----------------------------------------------------------------------------------------------------------------------------------------

# 18. L2.1 POWER CORD--------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL %in% "L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE == "POWER CORD",
         Category:= "L2.1-POWER CORD"]
raw.dt_d[Category=="L2.1-POWER CORD", Duty:="Power RD"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 19. L2.1 SMALL BOARD-------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("IO BOARD","POWER BOARD","SMALL BOARD","USB BOARD","VGA BOARD"),
         Category:="L2.1-SMALL BOARD"]
raw.dt_d[Category=="L2.1-SMALL BOARD", Duty:="EE RD"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 20. L2.1 Memory------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("MEMORY-IC","MEMORY"),
         Category:="L2.1-MEMORY"]
raw.dt_d[Category=="L2.1-MEMORY", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------

#----------------------------------------------------------------------------------------------------------------------------------------

# 22. L2.1 RESISTOR------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("BRACKET","BASEBOARD","RESISTOR","PROGRAMABLE","IC","FLASH","CONN","CABLE","PACKING","CAPACITOR","INDUCTOR","DISCRETE","SHIELDING"),
         Category:="L2.1-MB Material"]
raw.dt_d[Category=="L2.1-MB Material", Duty:="EE RD"]
#----------------------------------------------------------------------------------------------------------------------------------------
# 22. L2.1 BASEBOARD------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("BASEBOARD"),
         Category:="L2.1-BASEBOARD"] 
raw.dt_d[Category=="L2.1-BASEBOARD", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 23. L2.1 RESISTOR----------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("SSD"),
         Category:="L2.1-SSD"]
raw.dt_d[Category=="L2.1-SSD", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 24. L2.1 Card Reader-------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("CARD READER"),
         Category:="L2.1-CARD READER"]
raw.dt_d[Category=="L2.1-CARD READER", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 25. L2.1 CPU---------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("CPU"),
         Category:="L2.1-CPU"]
raw.dt_d[Category=="L2.1-CPU", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 26. L2.1 ODD---------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("ODD"),
         Category:="L2.1-ODD"]
raw.dt_d[Category=="L2.1-ODD", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 27. L2.1 CARD READER--------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("CARD READER"),
         Category:="L2.1-CARD READER"]
raw.dt_d[Category=="L2.1-CARD READER", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 28. L2.1 THERMAL--------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("THERMAL","FAN"),
         Category:="L2.1-THERMAL"]
raw.dt_d[Category=="L2.1-THERMAL", Duty:="THERMAL RD"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 29. L2.1 GPU Card--------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("GPU CARD"),
         Category:="L2.1-GPU Card"]
raw.dt_d[Category=="L2.1-GPU Card", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------
# 29. L2.1 INSOURCING CARD--------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("INSOURCING CARD") & ORG_REPLACED_PART_DESC %like% "ASUS RX7900XT-20G-PASSIVE PCIE CARD",
         Category:="L2.1-GPU Card"]
raw.dt_d[Category=="L2.1-GPU Card", Duty:="RD/SQA"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("INSOURCING CARD") & ORG_REPLACED_PART_DESC %like% "PIKE",
         Category:="L2.1-Raid Card"]
raw.dt_d[Category=="L2.1-Raid Card", Duty:="RD/SQA"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("INSOURCING CARD") & ORG_REPLACED_PART_DESC %like% "PEI",
         Category:="L2.1-Lan Card"]
raw.dt_d[Category=="L2.1-Lan Card", Duty:="RD/SQA"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" & 
           ORG_REPLACED_PART_DESC %like% "PEI|LAN CARD",
         Category:="L2.1-Lan Card"]
raw.dt_d[Category=="L2.1-Lan Card", Duty:="RD/SQA"]


raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("INSOURCING CARD") & ORG_REPLACED_PART_DESC %like% "PEX",
         Category:="L2.1-SMALL BOARD"]
raw.dt_d[Category=="L2.1-SMALL BOARD", Duty:="RD/SQA"]
# 29. L2.1 OUTSOURCING CARD--------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("OUTSOURCING CARD") & ORG_REPLACED_PART_DESC %like% c("CONNECTX-6", "LAN"),
         Category:="L2.1-Lan Card"]
raw.dt_d[Category=="L2.1-Lan Card", Duty:="RD/SQA"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("OUTSOURCING CARD") & ORG_REPLACED_PART_DESC %like% c("RAID"),
         Category:="L2.1-Raid Card"]
raw.dt_d[Category=="L2.1-Raid Card", Duty:="RD/SQA"]

# 22. L2.1 BASEBOARD PCBA------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL %in% "L2.1" & !grepl("^N", VIRTUAL_REASON_CODE) & 
           MUC_MODULE %in% c("GPU CARD") & grepl("PCBA", ORG_REPLACED_PART_DESC),
         Category := "L2.1-BASEBOARD PCBA"]
raw.dt_d[Category=="L2.1-BASEBOARD PCBA", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 30. L2.1 Card--------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("CARD")& 
           ORG_REPLACED_PART_DESC %like% "PIKE", 
         Category := "L2.1-Raid Card"]
raw.dt_d[Category == "L2.1-Raid Card", Duty := "RD/SQA"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("CARD")& 
           ORG_REPLACED_PART_DESC %like% "PEI", 
         Category := "L2.1-Lan Card"]
raw.dt_d[Category == "L2.1-Lan Card", Duty := "RD/SQA"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("CARD")& 
           ORG_REPLACED_PART_DESC %like% "PEX", 
         Category := "L2.1-SMALL BOARD"]
raw.dt_d[Category == "L2.1-SMALL BOARD", Duty := "EE RD"]

raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("CARD")& 
           ORG_REPLACED_PART_DESC %like% "ESC8000-SKU-PLX8796", 
         Category := "L2.1-SMALL BOARD"]
raw.dt_d[Category == "L2.1-SMALL BOARD", Duty := "EE RD"]
#----------------------------------------------------------------------------------------------------------------------------------------
# 31. L2.1 Card--------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & !VIRTUAL_REASON_CODE%like%"^N" &
           MUC_MODULE%in%c("CHIPSET","CHIPSET(BGA)"),
         Category:="L2.1-CHIPSET"]
raw.dt_d[Category=="L2.1-CHIPSET", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 32. L2.1 KP----------------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL%in%"L2.1" & Category=="OTHERS", 
         Category:="L2.1-Others RP Parts"]
raw.dt_d[Category=="L2.1-Others RP Parts", Duty:="RD/SQA"]
#----------------------------------------------------------------------------------------------------------------------------------------


# 33. L2.2 Reassembly--------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL=="L2.2" & !VIRTUAL_REASON_CODE %in% c("R14","R25"), Category:= "L2.2 Reassembly"]
raw.dt_d[Category=="L2.2 Reassembly", Duty:="EMS/SQE"]
#----------------------------------------------------------------------------------------------------------------------------------------

# 34. 非品質相關問題---------------------------------------------------------------------------------------------------------------------
raw.dt_d[REPAIR_LEVEL=="L2.2" & VIRTUAL_REASON_CODE%in%c("R14","R25") & !CUSTOMER_COUNTRY_CODE_DESC=="CHINA", Category:="Not a Quality Issue"]
raw.dt_d[EXCEPTION_SUB_REASON%in%"X34-2"&RMA_TYPE_REASON%in%"REAS-29", Category:="Not a quality issue"]
raw.dt_d[Category=="L0 DOA Others"                     & VIRTUAL_REASON_CODE=="F16" & REASON_SUB_DESC%like%"CID", Category:="Not a quality issue"]
raw.dt_d[Category=="L0 RMA Others (NTF, Customer decides not to repair)" & VIRTUAL_REASON_CODE=="F16" & REASON_SUB_DESC %like% "CID", Category:="Not a Quality Issue"]
raw.dt_d[Category=="L0 DOA Others" & PROBLEM%in%c("",NA), Category:="Not a quality issue"]
raw.dt_d[Category=="L2.2 Reassembly" & VIRTUAL_REASON_CODE%in%c("F19","F22","R23"), Category:="Not a Quality Issue"]
raw.dt_d[Category=="Not a Quality Issue", Duty:="Not a Quality Issue"]
#--------------------------------------------------------------------------------------------------------------------------------------------------------

# 35. 部份歸類修正---------------------------------------------------------------------------------------------------------------------------------------
raw.dt_d[Category=="L0 DOA Others" & CUSTOMER_COUNTRY_CODE_DESC=="CHINA" & R16_F16_DOA =="Y", Category:="L0 DOA Credit w/ problem"]
raw.dt_d[Category=="L0 RMA Others (NTF, Customer decides not to repair.)" & CUSTOMER_COUNTRY_CODE_DESC=="CHINA" & R16_F16_DOA == "Y", Category:="L0 DOA Credit w/ problem"]

raw.dt_d[Category=="L0 DOA Credit w/ problem", Duty:="EMS/SQE"]

#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
char_cols <- names(raw.dt_d)[sapply(raw.dt_d, is.character)]
raw.dt_d[, (char_cols) := lapply(.SD, function(x) { 
  x[x == ""] <- NA
  return(x)
}), .SDcols = char_cols]

# 刪除完全空白的列（所有欄位皆為 NA 的列）
DataClean_raw.dt_d <- raw.dt_d[, which(colSums(!is.na(raw.dt_d)) > 0), with = FALSE]

write_feather(DataClean_raw.dt_d, "C:/Users/rex4_chen/Desktop/feather/DataClean_raw.dt_d.feather")

write.xlsx(DataClean_raw.dt_d, "C:/Users/rex4_chen/Desktop/DataClean_raw.dt_20250611.xlsx")
#DataClean_raw.dt_d$`PART_DESC(MODEL)`[DataClean_raw.dt_d$ORG_MODEL_PART_NO == "90SF02T7-M01W10"]

 