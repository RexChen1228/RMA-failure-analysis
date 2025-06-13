DataClean_raw.dt_d <- read_feather("C:/Users/rex4_chen/Desktop/feather/DataClean_raw.dt_d.feather")
#先篩選條件
#做分析的時候不要排除一筆多修和CID/OOW
DataClean_raw.dt_d <- DataClean_raw.dt_d[REPAIR_LEVEL != "CID/OOW",]
DataClean_raw.dt_d <- DataClean_raw.dt_d[grepl("^(R|S|T)", SN2)]
DataClean_raw.dt_d <- DataClean_raw.dt_d[Mix_Repair_LMD == "Main_Repair",]

L0L2 <- DataClean_raw.dt_d[FEBE_N %like% ('FE'),]
DataClean_SYS <- L0L2[PRODUCT_LINE %in% c("S","SV","SF"),]

DataClean_SYS <- DataClean_SYS[TYPE == "System",]

#L6L10
DataClean_L6 <- DataClean_SYS[L6_L10 == "準系統",]
DataClean_L10 <- DataClean_SYS[L6_L10 == "全系統",]
#SB
DataClean_SB <- L0L2[PRODUCT_LINE %in% "SB",]
DataClean_SB <- DataClean_SB[ODM_OEM == "channel",]
#L3L4
L3L4 <- DataClean_raw.dt_d[FEBE_N %like% ('BE'),]
L3L4 <- L3L4[TYPE == "Board",]
L3L4 <- L3L4[ODM_OEM == "channel",]

# 創建一個新的工作簿
L0L2_L3L4Rawdata <- createWorkbook()
# 將第 1 個工作表添加到新文件
addWorksheet(L0L2_L3L4Rawdata, "L6L10")
writeData(L0L2_L3L4Rawdata, "L6L10", DataClean_SYS)

addWorksheet(L0L2_L3L4Rawdata, "L6")
writeData(L0L2_L3L4Rawdata, "L6", DataClean_L6)

addWorksheet(L0L2_L3L4Rawdata, "L10")
writeData(L0L2_L3L4Rawdata, "L10", DataClean_L10)

addWorksheet(L0L2_L3L4Rawdata, "SB")
writeData(L0L2_L3L4Rawdata, "SB", DataClean_SB)

addWorksheet(L0L2_L3L4Rawdata, "L3L4")
writeData(L0L2_L3L4Rawdata, "L3L4", L3L4)

today_date <- format(Sys.Date(), "%Y%m%d")

file_path <- paste0("C:/Users/rex4_chen/Desktop/L0L2_L3L4Rawdata", today_date, ".xlsx")
# 保存 Excel 文件
saveWorkbook(L0L2_L3L4Rawdata, file_path, overwrite = TRUE)


