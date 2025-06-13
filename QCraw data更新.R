
old_data_file <- "C:/Users/rex4_chen/Desktop/feather/QCRawData_All.feather"


# 自動生成前一個月的 YYYYMM 格式
last_month <- format(seq.Date(as.Date(format(Sys.Date(), "%Y-%m-01")), by = "-1 month", length = 2)[2], "%Y%m")

new_file_path <- paste0("C:/Users/rex4_chen/Desktop/QCRawData-", last_month, ".xlsx")

if (file.exists(old_data_file)) {
  old_data <- read_feather(old_data_file)
} else {
  old_data <- data.table() # 若沒有舊數據文件，創建一個空的data.table
}

new_data <- read_excel(new_file_path)
new_data <- data.table(new_data)

QCRawData_All <- rbind(old_data, new_data, fill = TRUE)

write_feather(QCRawData_All, old_data_file)
