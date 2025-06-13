
service_level <- function(DataClean, shipment25, shipment24){
  
  sum_shipment <- function(data, months) {
    sum(data$Total_qtr[data$Month %in% months], na.rm = TRUE)
  }
  #須改
  sums_2024 <- sum_shipment(shipment24 ,1:5)
  sums_2025 <- sum_shipment(shipment25 ,1:5)
  
  #須改
  service_level25 <- subset(DataClean, grepl("^2025/(0[1-5])$", Close_YM))
  service_level24 <- subset(DataClean, grepl("^2024/(0[1-5])$", Close_YM))
  
  service_level25 <- subset(service_level25, SN2_Y == 2025)
  service_level24 <- subset(service_level24, SN2_Y == 2024)
  
  service_level25 <- dcast(service_level25, Category + Duty  ~ SN2_Y, value.var = "SN_RN", fun.aggregate = length)
  service_level24 <- dcast(service_level24, Category + Duty  ~ SN2_Y, value.var = "SN_RN", fun.aggregate = length)
  
  service_level <- full_join(service_level25, service_level24, by = c("Category", "Duty"))
  service_level <- service_level[order(service_level$`2025` ,decreasing = TRUE), ]
  
  service_level <- add_row(service_level, Category = "Total", Duty = "", `2025` = sum(service_level$`2025`, na.rm = T), `2024` = sum(service_level$`2024`, na.rm = T))
  
  service2025 <- sprintf("%.2f%%", (service_level[[3]] / rep(service_level[nrow(service_level), 3][[1]], nrow(service_level))) * 100)
  service2024 <- sprintf("%.2f%%", (service_level[[4]] / rep(service_level[nrow(service_level), 4][[1]], nrow(service_level))) * 100)
  
  service_level <- cbind(service_level, service2025, service2024, sums_2025[1], sums_2024[1])
  
  CRR25 <- sprintf("%.2f%%", (service_level[[3]] / sums_2025[1])* 100)
  CRR24 <- sprintf("%.2f%%", (service_level[[4]] / sums_2024[1])* 100)
  
  Improvement_rate <- sprintf(
    "%.2f%%", 
    (
      (
        (service_level[[3]] / sums_2024[1]) - 
          (service_level[[4]] / sums_2025[1])
      ) / 
        (service_level[[4]] / sums_2024[1])
    ) * 100
  )
  
  gap <- sprintf("%.2f%%", ((service_level[[3]] / sums_2025[1])-(service_level[[4]] / sums_2024[1]))*100)
  
  service_level <- cbind(service_level, CRR25, CRR24, Improvement_rate, gap)
  
  service_level <- service_level[, c(1, 3, 5, 9, 4, 6, 10, 11, 12, 2, 7, 8)]
  
  # 獲取當前月份，並轉換為數字，然後減一
  previous_month <- as.numeric(format(Sys.Date(), "%m")) - 1 
  
  # 動態生成欄位名稱
  column_name <- paste0("Category-LT", previous_month)
  
  
  colnames(service_level) <- c(column_name, "2025Qty", "2025service", "2025CRR",
                               "2024Qty", "2024service", "2024CRR",  "Improvement_rate", "gap", "owner", "25Ship", "24Ship")
  
  service_level[-1,11:12] <- NA
  return(service_level)
}

L6L10_service_level <- service_level(DataClean_SYS, SYSshipment25, SYSshipment24)
L6_service_level <- service_level(DataClean_L6, L6shipment25, L6shipment24)
L10_service_level <- service_level(DataClean_L10, L10shipment25, L10shipment24)
SB_service_level <-service_level(DataClean_SB, SBshipment25, SBshipment24)
# 讀取原始 Excel 文件的各個工作表
file_path <- "C:/Users/rex4_chen/Desktop/Servicelevel20241210.xlsx"
Servicelevel <- loadWorkbook(file_path)
all_sheets <- getSheetNames(file_path)
history <- all_sheets[1]

# 刪除其他工作表
for (sheet in all_sheets[-1]) {
  removeWorksheet(Servicelevel, sheet)
}

# 讀取第一個工作表的資料，這個工作表不會改動
history <- read.xlsx(Servicelevel, sheet = 1)

addWorksheet(Servicelevel, "L6L10")
writeData(Servicelevel, "L6L10", L6L10_service_level)

addWorksheet(Servicelevel, "L10")
writeData(Servicelevel, "L10", L10_service_level)

addWorksheet(Servicelevel, "L6")
writeData(Servicelevel, "L6", L6_service_level)

addWorksheet(Servicelevel, "SB")
writeData(Servicelevel, "SB", SB_service_level)
# 保存 Excel 文件
saveWorkbook(Servicelevel, "C:/Users/rex4_chen/Desktop/Servicelevel new 20250611.xlsx", overwrite = TRUE)

