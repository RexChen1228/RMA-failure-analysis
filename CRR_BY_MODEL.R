DataClean_raw.dt_d <- read_feather("C:/Users/rex4_chen/Desktop/feather/DataClean_raw.dt_d.feather")

#做分析的時候不要排除一筆多修理
Data_SYS <- DataClean_raw.dt_d[REPAIR_LEVEL != "CID/OOW",]
Data_SYS <- Data_SYS[FEBE_N %like% ('FE'),]
Data_SYS <- Data_SYS[PRODUCT_LINE %in% c("S","SV","SF"),]
Data_SYS <- Data_SYS[TYPE == "System",]
#先篩選條件
DataClean_raw.dt_d <- DataClean_raw.dt_d[REPAIR_LEVEL != "CID/OOW",]
L0L2 <- DataClean_raw.dt_d[FEBE_N %like% ('FE'),]
L0L2 <- L0L2[Mix_Repair_LMD == "Main_Repair",]
DataClean_SYS <- L0L2[PRODUCT_LINE %in% c("S","SV","SF"),]
DataClean_SYS <- DataClean_SYS[TYPE == "System",]
 
repair_data <- dcast(DataClean_SYS, MODEL_NAME ~ Close_YM, value.var = "SN_RN", fun.aggregate = length)

ShipmentData_ALL <- read_feather("C:/Users/rex4_chen/Desktop/feather/ShipmentData_ALL.feather")

ori_raw.dt <- data.table(ShipmentData_ALL)
raw.dt_o <- ori_raw.dt[!duplicated(ori_raw.dt),]

raw.dt_o <- raw.dt_o[grepl("Server|ISG",BU)]
raw.dt_o <- raw.dt_o[ `Product Line` %in% c("SF", "SV","S")]
raw.dt_PRO <- raw.dt_o[!(grepl("E500|E800|E900|^EG", Model))]

WS_Mapping <-read_excel("D:/PQA/Raw_Data/Mapping/Workstation_list.xlsx")
setDT(WS_Mapping)

raw.dt_PRO <- raw.dt_PRO[`Model` %in% WS_Mapping$Model, WS := "Y"]

raw.dt_PRO <- raw.dt_PRO[is.na(WS), WS := "N"]

# 只保留 WS 為 "N" 的行
raw.dt_PRO <- raw.dt_PRO[WS == "N",]

raw.dt_PRO[, Month := as.character(Month)]
raw.dt_PRO[, Year := as.character(Year)]
raw.dt_PRO[, Ship_YM := paste(Year, sprintf("%02d", as.numeric(Month)), sep = "/")]
raw.dt_PRO$model_new <- gsub("(.*?)(/.*)?", "\\1", raw.dt_PRO$`Product Desc`)
raw.dt_PRO[, model_new := gsub("\\s*\\(.*?\\)|\\s*<.*?>", "", model_new)]



raw.dt_PRO[, ORG_MODEL_PART_DESC := ifelse(grepl("//", `Product Desc`),
                                   sub("//.*", "", `Product Desc`),
                                   `Product Desc`)]

model_mapping <- DataClean_SYS[, .SD[1], by = ORG_MODEL_PART_DESC]

# 左表為主，做 left join
raw.dt_PRO <- merge(
  raw.dt_PRO,
  model_mapping[, .(ORG_MODEL_PART_DESC, MODEL_NAME)],
  by = "ORG_MODEL_PART_DESC",
  all.x = TRUE
)


# 新增欄位：優先用維修 model_name，否則用出貨 model_new
raw.dt_PRO[, model_name_final := fcoalesce(MODEL_NAME, model_new)]


shipment_data <- dcast(raw.dt_PRO, model_name_final ~ Ship_YM, value.var = 'Net QTY', fun.aggregate = sum)
#資料怪怪的先刪掉12m再來改
#shipment_data <- shipment_data[, 1:(ncol(shipment_data) - 1)]

setnames(shipment_data, "model_name_final", "MODEL_NAME")
# 將資料轉換為長格式
shipment_long <- data.table::melt(shipment_data, id.vars = "MODEL_NAME", variable.name = "Date", value.name = "Shipment_QTY")
repair_long <- data.table::melt(repair_data, id.vars = "MODEL_NAME", variable.name = "Date", value.name = "Repair_QTY")

# 合併出貨和返修資料
combined_data <- merge(shipment_long, repair_long, by = c("MODEL_NAME", "Date"), all.x = TRUE)

# 將 NA 的 Repair_QTY 填充為 0
combined_data[is.na(Repair_QTY), Repair_QTY := 0]

# 計算累積出貨和累積返修數量
combined_data[, Cumulative_Shipment := cumsum(Shipment_QTY), by = MODEL_NAME]
combined_data[, Cumulative_Repair := cumsum(Repair_QTY), by = MODEL_NAME]

# 計算 Repair_Ratio 並格式化為百分比小數點後兩位
combined_data[, Repair_Ratio := sprintf("%.2f%%", (Cumulative_Repair / Cumulative_Shipment) * 100)]

# 去掉包含 NA 值的行
clean_data <- combined_data[Repair_Ratio != "NaN%"]

# 使用 dcast 將長格式轉換為寬格式
wide_data <- dcast(clean_data, MODEL_NAME ~ Date, value.var = "Repair_Ratio")
wide_data <- wide_data[MODEL_NAME != "ESC4000",]

dt <- copy(wide_data)  # 建立資料副本，以免覆蓋原始資料

shift_left_all <- function(row) {
  non_na_values <- row[!is.na(row)]  # 提取非 NA 的值
  c(non_na_values, rep(NA, length(row) - length(non_na_values)))  # 左對齊，右側填充 NA
}

# 對每一行應用 shift_left_all 函數
dt[, paste0("LT", 1:(ncol(dt) - 1)) := as.list(shift_left_all(unlist(.SD))), by = 1:nrow(dt), .SDcols = names(dt)[-1]]

lt_columns <- grep("^LT", names(dt), value = TRUE)  # 查找所有以 "LT" 開頭的現有列

# 只保留 MODEL_NAME 和現有的 LT 列
dt <- dt[, c("MODEL_NAME", lt_columns), with = FALSE]
# 新增一列 LT，計算每行非 NA 值的個數
dt[, LT := rowSums(!is.na(.SD)), .SDcols = patterns("^LT")]

# 重新排列列順序，使得 LT 列位於第二列
setcolorder(dt, c("MODEL_NAME", "LT", setdiff(names(dt), c("MODEL_NAME", "LT"))))

CRR_BY_MODEL <- dt[order(LT), ]

Shipmentdt <- dcast(clean_data, MODEL_NAME ~ Date, value.var = "Shipment_QTY")

Shipmentdt[, paste0("LT", 1:(ncol(Shipmentdt) - 1)) := as.list(shift_left_all(unlist(.SD))), by = 1:nrow(Shipmentdt), .SDcols = names(Shipmentdt)[-1]]

Shipmentdt <- Shipmentdt[, c("MODEL_NAME", lt_columns), with = FALSE]

Shipmentdt[, LT := rowSums(!is.na(.SD)), .SDcols = patterns("^LT")]

setcolorder(Shipmentdt, c("MODEL_NAME", "LT", setdiff(names(Shipmentdt), c("MODEL_NAME", "LT"))))

Shipmentdt <- Shipmentdt[order(LT), ]

Repairdt <- dcast(clean_data, MODEL_NAME ~ Date, value.var = "Repair_QTY")

Repairdt[, paste0("LT", 1:(ncol(Repairdt) - 1)) := as.list(shift_left_all(unlist(.SD))), by = 1:nrow(Repairdt), .SDcols = names(Repairdt)[-1]]

Repairdt <- Repairdt[, c("MODEL_NAME", lt_columns), with = FALSE]

Repairdt[, LT := rowSums(!is.na(.SD)), .SDcols = patterns("^LT")]

setcolorder(Repairdt, c("MODEL_NAME", "LT", setdiff(names(dt), c("MODEL_NAME", "LT"))))

Repairdt <- Repairdt[order(LT), ]

# 讀取原始 Excel 文件的各個工作表
file_path <- "C:/Users/rex4_chen/Desktop/CRR_BY_MODEL_241203.xlsx"
CRRBYMODEL <- loadWorkbook(file_path)
all_sheets <- getSheetNames(file_path)
history <- all_sheets[1]

# 刪除其他工作表
for (sheet in all_sheets[-1]) {
  removeWorksheet(CRRBYMODEL, sheet)
}

# 讀取第一個工作表的資料，這個工作表不會改動
history <- read.xlsx(CRRBYMODEL, sheet = 1)

addWorksheet(CRRBYMODEL, "CRR")
writeData(CRRBYMODEL, "CRR", CRR_BY_MODEL)

addWorksheet(CRRBYMODEL, "shipment")
writeData(CRRBYMODEL, "shipment", Shipmentdt)

addWorksheet(CRRBYMODEL, "return")
writeData(CRRBYMODEL, "return", Repairdt)

addWorksheet(CRRBYMODEL, "raw data")
writeData(CRRBYMODEL, "raw data", Data_SYS)

today_date <- format(Sys.Date(), "%Y%m%d")

file_path <- paste0("C:/Users/rex4_chen/Desktop/CRR_BY_MODEL_", today_date, ".xlsx")

# 保存 Excel 文件
saveWorkbook(CRRBYMODEL, file_path, overwrite = TRUE)



