# 定義來源資料夾和儲存資料夾
source_folder <- "C:/Users/rex4_chen/Desktop/出貨原始資料/"

target_folder <- "C:/Users/rex4_chen/Desktop/feather/"

# 取得所有符合 "DL-98_BIS數據協助" 開頭的 Excel 文件
excel_files <- list.files(path = source_folder, pattern = "^DL-98_BIS數據協助.*\\.xlsx$", full.names = TRUE)

# 迴圈處理每個 Excel 文件
for (file in excel_files) {
  # 生成 Feather 文件名
  feather_file <- gsub(".xlsx$", ".feather", file)
  feather_file <- file.path(target_folder, basename(feather_file))
  
  # 檢查 Feather 文件是否存在，且 Excel 文件有無更新
  if (!file.exists(feather_file) || file.info(file)$mtime > file.info(feather_file)$mtime) {
    # 讀取 Excel 檔案
    df <- read_excel(file)
    
    # 儲存為 Feather 格式
    write_feather(df, feather_file)
    cat("轉換:", feather_file, "\n")
  } else {
    cat("跳過已存在的文件:", feather_file, "\n")
  }
}


# 取得所有 Feather 檔案
feather_files <- list.files(path = target_folder, pattern = "^DL-98_BIS數據協助.*\\.feather$", full.names = TRUE)

# 讀取並合併所有 Feather 檔案
data.ls <- lapply(feather_files, read_feather)
ShipmentData_ALL <- rbindlist(data.ls, fill = TRUE)  # 合併成一個 data.table
old_data_file <- "C:/Users/rex4_chen/Desktop/feather/ShipmentData_ALL.feather"
write_feather(ShipmentData_ALL, old_data_file)












