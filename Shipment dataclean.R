ShipmentData_ALL <- read_feather("C:/Users/rex4_chen/Desktop/feather/ShipmentData_ALL.feather")

ori_raw.dt <- data.table(ShipmentData_ALL)
raw.dt_o <- ori_raw.dt[!duplicated(ori_raw.dt),]

raw.dt_o <- raw.dt_o[grepl("Server|ISG",BU)]
raw.dt_PRO <- raw.dt_o[!(grepl("E500|E800|E900|^EG", Model))]
WS_Mapping <-read_excel("D:/PQA/Raw_Data/Mapping/Workstation_list.xlsx")

setDT(WS_Mapping)

raw.dt_PRO <- raw.dt_PRO[`Model` %in% WS_Mapping$Model, WS := "Y"]
raw.dt_PRO <- raw.dt_PRO[is.na(WS), WS := "N"]
raw.dt_PRO <- raw.dt_PRO[WS == "N",]

Server_Level <- read_excel("D:/PQA/Raw_Data/Mapping/Mapping_Table_SYS_Level.xlsx", sheet = "SYS_Level")

raw.dt_PRO[ ,L6_L10 := Server_Level$SYS_Level[match(raw.dt_PRO$'Part Number',Server_Level$Part_Number)]]
raw.dt_PRO[`Product Line`%in% c("S","SV","SF") & `Product Desc`%like% c("WOCPU"), L6_L10 := "準系統"]
raw.dt_PRO[`Product Line` %in% c("S","SV","SF") & is.na(L6_L10), L6_L10 := "全系統"]

raw.dt_PRO[, Month := as.character(Month)]
raw.dt_PRO[, Year := as.character(Year)]
raw.dt_PRO[, Ship_YM := paste(Year, sprintf("%02d", as.numeric(Month)), sep = "/")]
raw.dt_PRO$model_new <- gsub("(.*?)(/.*)?", "\\1", raw.dt_PRO$`Product Desc`)

filter_and_summarize <- function(data, product_lines, year = NULL, L6_L10_value = NULL) {
  filtered_data <- data[`Product Line` %in% product_lines]
  
  # 可選年份篩選
  if (!is.null(year)) {
    filtered_data <- filtered_data[Year == year]
  }
  
  # 可選 L6_L10 篩選
  if (!is.null(L6_L10_value)) {
    filtered_data <- filtered_data[L6_L10 == L6_L10_value]
  }
  
  # 匯總
  summarized_data <- filtered_data[, .(Total_qtr = sum(`Net QTY`, na.rm = TRUE)), by = .(Year, Month)]
  summarized_data <- summarized_data[order(Year, as.numeric(Month))]
  
  return(summarized_data)
}

# 呼叫篩選與總結函數
SYSshipment23 <- filter_and_summarize(raw.dt_PRO, c("SF", "SV", "S"), year = "2023")
SYSshipment24 <- filter_and_summarize(raw.dt_PRO, c("SF", "SV", "S"), year = "2024")
SYSshipment25 <- filter_and_summarize(raw.dt_PRO, c("SF", "SV", "S"), year = "2025")

L6shipment23 <- filter_and_summarize(raw.dt_PRO, c("SF", "SV", "S"), year = "2023", L6_L10 = "準系統")
L6shipment24 <- filter_and_summarize(raw.dt_PRO, c("SF", "SV", "S"), year = "2024", L6_L10 = "準系統")
L6shipment25 <- filter_and_summarize(raw.dt_PRO, c("SF", "SV", "S"), year = "2025", L6_L10 = "準系統")

L10shipment23 <- filter_and_summarize(raw.dt_PRO, c("SF", "SV", "S"), year = "2023", L6_L10 = "全系統")
L10shipment24 <- filter_and_summarize(raw.dt_PRO, c("SF", "SV", "S"), year = "2024", L6_L10 = "全系統")
L10shipment25 <- filter_and_summarize(raw.dt_PRO, c("SF", "SV", "S"), year = "2025", L6_L10 = "全系統")

SBshipment23 <- filter_and_summarize(raw.dt_PRO, c("SB"), year = "2023")
SBshipment24 <- filter_and_summarize(raw.dt_PRO, c("SB"), year = "2024")
SBshipment25 <- filter_and_summarize(raw.dt_PRO, c("SB"), year = "2025")

#需要的話自己生產
#
file_path <- paste0("C:/Users/rex4_chen/Desktop/Shipment20250611.xlsx")
# 產出Excel檔
#
write.xlsx(raw.dt_PRO, file_path,  overwrite = TRUE)

