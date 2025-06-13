adjust_table <- function(dt) {
  
  # 確定需要操作的列（去掉 "Year" 列）
  cols <- setdiff(names(dt), "Year")
  
  # 遍歷每一行
  for (i in seq_len(nrow(dt))) {
    year <- as.character(dt$Year[i])  # 提取 Year 列的值
    
    # 遍歷所有的列，根據年份來決定是否保留資料
    dt[i, (cols) := lapply(cols, function(col) {
      # 比較年份，若列名中的年份大於等於該行的 "Year"，則保留原值，否則設為 NA
      if (as.numeric(substr(col, 1, 4)) >= as.numeric(year)) {
        return(dt[[col]][i])  # 保留原值
      } else {
        return(NA)  # 設為 NA
      }
    })]
  }
  
  return(dt)
}
shift_left_all <- function(row) {
  non_na_values <- row[!is.na(row)]  # 提取非 NA 的值
  c(non_na_values, rep(NA, length(row) - length(non_na_values)))  # 左對齊，右側填充 NA
}
#有資料
data_table_and_plot <- function(cleandata,shipdata23,shipdata24,shipdata25) {
  #將返修資料存起來
  pivot_table1 <- dcast(cleandata, Close_YM ~ SN2_Y, value.var = "SN_RN", fun.aggregate = length)
  pivot_table1 <- as.data.table(pivot_table1)
  #創建空白且都是0的表格
  years <- c(2023, 2024, 2025)
  months <- sprintf("%02d", 1:12)  
  # 創建所有組合
  date_combinations <- CJ(Year = years, Month = months)
  
  # 轉換成 "YYYY/MM" 格式作為索引
  date_combinations[, Close_YM := paste(Year, Month, sep = "/")]
  
  # 建立最終的資料表，並填充 0
  final_table <- dcast(date_combinations, Close_YM ~ Year, value.var = "Year", fun.aggregate = function(x) 0)
  
  # 使用 left_join 進行合併（如果尚未合併）
  merged_table <- left_join(final_table, pivot_table1, by = "Close_YM")
  
  # 只保留 2023.y、2024.y、2025，並改名稱
  merged_table <- merged_table %>%
    select(Close_YM, `2023.y`, `2024.y`, `2025.y`) %>%
    rename(`2023` = `2023.y`, `2024` = `2024.y`, `2025` = `2025.y`) %>%
    # 用 0 填補所有 NA 值
    mutate(across(c(`2023`, `2024`, `2025`), ~replace_na(., 0)))

  # 將資料從長格式轉為寬格式
  dt_long1 <- data.table::melt(merged_table, id.vars = "Close_YM", variable.name = "Year", value.name = "return")
  
  returnqty1 <- data.table::dcast(dt_long1, Year ~ Close_YM, value.var = "return")
  
  adjust_table(returnqty1)
  # 對每一行應用 shift_left_all 函數
  returnqty1[, paste0("LT", 1:(ncol(returnqty1) - 1)) := as.list(shift_left_all(unlist(.SD))), by = 1:nrow(returnqty1), .SDcols = names(returnqty1)[-1]]
  
  lt_columns1 <- grep("^LT", names(returnqty1), value = TRUE)  # 查找所有以 "LT" 開頭的現有列
  returnqty1 <- returnqty1[, c("Year", lt_columns1), with = FALSE]
  
  returnqty1$Year <- as.character(returnqty1$Year)
  
  returnqty <- copy(returnqty1)
  
  cumsum_1_row <- cumsum(unlist(returnqty[1, -1]))
  cumsum_2_row <- cumsum(unlist(returnqty[2, -1]))
  cumsum_3_row <- cumsum(unlist(returnqty[3, -1]))
  cusqty <- rbind(cumsum_1_row,cumsum_2_row,cumsum_3_row)
  
  returnqty[,1] <- c("2023SN", "2024SN", "2025SN")
  returnqty <- cbind(returnqty[,1],cusqty)
  
  #出貨資料處理
  ship23qty <- cumsum(shipdata23$Total_qtr)
  ship24qty <- cumsum(shipdata24$Total_qtr)
  ship25qty <- cumsum(shipdata25$Total_qtr)
  
  # 補齊12個元素，缺少的用0補充
  length(ship25qty) <- 12  
  # 如果有空位，則用0填充
  ship25qty[is.na(ship24qty)] <- 0
  
  # 合併
  shipment <- rbind(ship23qty,ship24qty, ship25qty)
  shipmentqty <- data.table(shipment)
  
  # 找出長度
  len_wide_dt <- length( returnqty[,-1])
  len_shipment <- length(shipmentqty)
  
  # 如果 b 資料< a 資料，將 b 調整為與 a 長度相同
  if(len_shipment < len_wide_dt) {
    shipmentqty1 <- rep(shipmentqty[[ncol(shipmentqty)]],each = len_wide_dt-len_shipment)  # 用 b 的最後一個值填充至 a 的長度
  }
  
  # 轉換成矩陣，每列是一個值
  result <- matrix(shipmentqty1, nrow = length(shipmentqty[[length(shipmentqty)]]), byrow = TRUE)
  
  shipmentqty <- cbind(shipmentqty,result)
  
  CRRdecimal <- returnqty[,-1]/shipmentqty
  
  #出貨資料處理
  ship23qty1 <- shipdata23$Total_qtr
  ship24qty1 <- shipdata24$Total_qtr
  ship25qty1 <- shipdata25$Total_qtr
  # 補齊12個元素，缺少的用0補充
  length(ship25qty1) <- 12  
  # 如果有空位，則用0填充
  ship25qty1[is.na(ship25qty1)] <- 0
  
  # 合併
  shipment1 <- rbind(ship23qty1,ship24qty1, ship25qty1)
  shipmentqty1 <- data.table(shipment1)
  
  # 找出長度
  len_wide_dt1 <- length( returnqty1[,-1])
  len_shipment1 <- length(shipmentqty1)
  
  # 如果 b 資料< a 資料，將 b 調整為與 a 長度相同
  if(len_shipment1 < len_wide_dt1) {
    shipmentqty2 <- rep(shipmentqty1[[ncol(shipmentqty1)]],each = len_wide_dt1-len_shipment1)  # 用 b 的最後一個值填充至 a 的長度
  }
  
  # 轉換成矩陣，每列是一個值
  result1 <- matrix(shipmentqty2, nrow = length(shipmentqty1[[length(shipmentqty1)]]), byrow = TRUE)
  
  shipmentqty1 <- cbind(shipmentqty1,result1)
  
  
  df_percent <- CRRdecimal
  df_percent[] <- lapply(CRRdecimal, function(x) {
    # 判斷是否為小數點數字，並忽略 NA
    if (any(x %% 1 != 0, na.rm = TRUE)) {
      # 將小數轉換為百分比格式
      return(paste0(formatC(x * 100, format = "f", digits = 2), "%"))
    } else {
      # 直接返回原值
      return(x)
    }
  })
  
  CRR <- cbind(returnqty[,1],df_percent)
  
  # 使用 gather() 函數將數據轉換為長格式
  data_long <- gather(CRR, key = "Lifetime", value = "SN", -Year)
  
  # 移除百分號並將其轉換為數字格式
  data_long$SN <- as.numeric(gsub("%", "", data_long$SN)) / 100
  
  # 確保 Lifetime 列按照數字順序排列
  data_long$Lifetime <- factor(data_long$Lifetime, 
                               levels = paste0("LT", 1:(length(CRR)-1)), 
                               ordered = TRUE)
  setDT(data_long)
  df_clean <- data_long %>%
    filter(!is.na(Lifetime) & !is.na(SN) & !is.infinite(Lifetime) & !is.infinite(SN))
  # 計算每個 Year 的最後一個有效的 SN 值
  last_valid_SN <- df_clean[!is.na(SN), .(last_SN = SN[.N]), by = Year]
  
  # 設置顯示條件：六的倍數和每年最後的有效值
  df_clean$show_label <- ifelse(
    (as.numeric(gsub("LT", "", df_clean$Lifetime)) %% 6 == 0) |
      (df_clean$SN == last_valid_SN$last_SN[match(df_clean$Year, last_valid_SN$Year)]),
    TRUE, 
    FALSE
  )
  #每月須改
  df_clean_filtered <- df_clean %>%
    # 篩選 2025SN 且 Lifetime 為 LT1~
    filter(Year == "2025SN" & Lifetime == "LT5") %>%
    # 篩選 2024SN 且 Lifetime 為 LT1 ~
    bind_rows(
      df_clean %>%
        filter(Year == "2024SN" & between(as.integer(substr(Lifetime, 3, 4)), 1, 17))
    ) %>%
    # 篩選 2023SN 且 Lifetime 為 LT1 ~
    bind_rows(
      df_clean %>%
        filter(Year == "2023SN" & between(as.integer(substr(Lifetime, 3, 4)), 1, 29))
    )

  
  # 使用 ggplot2 繪製折線圖
  p <- ggplot(df_clean_filtered, aes(x = Lifetime, y = SN, color = Year, group = Year)) +
    geom_line(linewidth = 2.5, na.rm = TRUE) +  # 畫折線
    geom_point(size = 5, na.rm = TRUE) +  # 增加資料點大小
    geom_text(aes(label = ifelse(!is.na(SN) & show_label, paste0(round(SN * 100, 2), "%"), "")),
              vjust = ifelse(df_clean_filtered$Year == "2025SN", -3.5, ifelse(df_clean_filtered$Year == "2024SN", -1.2, -0.5)), hjust = 0.5, size = 9, show.legend = FALSE) +
    scale_color_manual(values = c("blue", "orange", "springgreen4")) +  # 更柔和的顏色搭配
    scale_y_continuous(labels = scales::percent_format(scale = 100)) +  # Y 軸顯示為百分比
    coord_cartesian(clip = "off") + 
    theme_minimal() +
    theme(
      axis.text.x = element_text(size = 16, hjust = 1, color = "black"),  # X 軸字體大小調整，無旋轉
      axis.text.y = element_text(size = 16, color = "black"),  # Y 軸字體大小調整，無粗體
      plot.margin = margin(30, 20, 20, 10),  # 圖表邊距
      legend.position = "bottom",  # 圖例位置
      legend.text = element_text(size = 14),  # 調整圖例文字大小
      legend.title = element_blank(),  # 移除圖例標題
      axis.title.y = element_blank(),  # 移除 Y 軸標題
      legend.background = element_rect(fill = "white", color = "black"),  # 設置圖例背景
      panel.grid.major = element_line(color = "gray", size = 0.5),  # 增加主網格線以提高可讀性
      panel.grid.minor = element_line(color = "gray", size = 0.2)  # 增加次網格線
    )
  
  #製作完整圖表
  result1 <- (CRRdecimal[1, ] - CRRdecimal[2, ]) / CRRdecimal[1, ]
  result2 <- (CRRdecimal[2, ] - CRRdecimal[3, ]) / CRRdecimal[2, ]
  # 3. 將小數轉回百分比
  result_percent1 <- paste0(round(result1 * 100, 2), "%")
  result_percent2 <- paste0(round(result2 * 100, 2), "%")
  result_percent1 <- c("24SN VS 23SN", result_percent1)
  result_percent2 <- c("25SN VS 24SN", result_percent2)
  result_percent1 <- as.data.table(result_percent1)
  result_percent2 <- as.data.table(result_percent2)
  result_percent1 <- transpose(result_percent1)
  result_percent2 <- transpose(result_percent2)
  
  returnqty1[,1] <- c("2023SN返修", "2024SN返修", "2025SN返修")
  shipmentqty1 <- cbind(c("2023SN出貨", "2024SN出貨", "2025SN出貨"),shipmentqty1)
  
  complete_table <- rbindlist(list(CRR, result_percent1, result_percent2, returnqty1, shipmentqty1[,1:13]), use.names = FALSE, fill = TRUE)
  # 將表格中的 "NA%" 替換成真正的 NA
  complete_table[complete_table == "NA%"] <- NA
  return(list(table = complete_table, plot = p))
}
#無資料
data_table_and_plot_n <- function(cleandata,shipdata23,shipdata24,shipdata25) {
  #將返修資料存起來
  pivot_table1 <- dcast(cleandata, Close_YM ~ SN2_Y, value.var = "SN_RN", fun.aggregate = length)
  pivot_table1 <- as.data.table(pivot_table1)
  #創建空白且都是0的表格
  years <- c(2023, 2024, 2025)
  months <- sprintf("%02d", 1:12)  # 生成 "01" 到 "12"
  
  # 創建所有組合
  date_combinations <- CJ(Year = years, Month = months)
  
  # 轉換成 "YYYY/MM" 格式作為索引
  date_combinations[, Close_YM := paste(Year, Month, sep = "/")]
  
  # 建立最終的資料表，並填充 0
  final_table <- dcast(date_combinations, Close_YM ~ Year, value.var = "Year", fun.aggregate = function(x) 0)
  
  # 使用 left_join 進行合併（如果尚未合併）
  merged_table <- left_join(final_table, pivot_table1, by = "Close_YM")
  
  # 只保留 2023.y、2024.y、2025，並改名稱
  merged_table <- merged_table %>%
    select(Close_YM, `2023.y`, `2024.y`, `2025`) %>%
    rename(`2023` = `2023.y`, `2024` = `2024.y`, `2025` = `2025`) %>%
    # 用 0 填補所有 NA 值
    mutate(across(c(`2023`, `2024`, `2025`), ~replace_na(., 0)))
  
  # 將資料從長格式轉為寬格式
  dt_long1 <- data.table::melt(merged_table, id.vars = "Close_YM", variable.name = "Year", value.name = "return")
  
  returnqty1 <- data.table::dcast(dt_long1, Year ~ Close_YM, value.var = "return")
  
  adjust_table(returnqty1)
  # 對每一行應用 shift_left_all 函數
  returnqty1[, paste0("LT", 1:(ncol(returnqty1) - 1)) := as.list(shift_left_all(unlist(.SD))), by = 1:nrow(returnqty1), .SDcols = names(returnqty1)[-1]]
  
  lt_columns1 <- grep("^LT", names(returnqty1), value = TRUE)  # 查找所有以 "LT" 開頭的現有列
  returnqty1 <- returnqty1[, c("Year", lt_columns1), with = FALSE]
  
  returnqty1$Year <- as.character(returnqty1$Year)
  
  returnqty <- copy(returnqty1)
  
  cumsum_1_row <- cumsum(unlist(returnqty[1, -1]))
  cumsum_2_row <- cumsum(unlist(returnqty[2, -1]))
  cumsum_3_row <- cumsum(unlist(returnqty[3, -1]))
  cusqty <- rbind(cumsum_1_row,cumsum_2_row,cumsum_3_row)
  
  returnqty[,1] <- c("2023SN", "2024SN", "2025SN")
  returnqty <- cbind(returnqty[,1],cusqty)
  
  #出貨資料處理
  ship23qty <- cumsum(shipdata23$Total_qtr)
  ship24qty <- cumsum(shipdata24$Total_qtr)
  ship25qty <- cumsum(shipdata25$Total_qtr)
  
  # 補齊12個元素，缺少的用0補充
  length(ship25qty) <- 12  
  # 如果有空位，則用0填充
  ship25qty[is.na(ship24qty)] <- 0
  
  # 合併
  shipment <- rbind(ship23qty,ship24qty, ship25qty)
  shipmentqty <- data.table(shipment)
  
  # 找出長度
  len_wide_dt <- length( returnqty[,-1])
  len_shipment <- length(shipmentqty)
  
  # 如果 b 資料< a 資料，將 b 調整為與 a 長度相同
  if(len_shipment < len_wide_dt) {
    shipmentqty1 <- rep(shipmentqty[[ncol(shipmentqty)]],each = len_wide_dt-len_shipment)  # 用 b 的最後一個值填充至 a 的長度
  }
  
  # 轉換成矩陣，每列是一個值
  result <- matrix(shipmentqty1, nrow = length(shipmentqty[[length(shipmentqty)]]), byrow = TRUE)
  
  shipmentqty <- cbind(shipmentqty,result)
  
  CRRdecimal <- returnqty[,-1]/shipmentqty
  
  #出貨資料處理
  ship23qty1 <- shipdata23$Total_qtr
  ship24qty1 <- shipdata24$Total_qtr
  ship25qty1 <- shipdata25$Total_qtr
  # 補齊12個元素，缺少的用0補充
  length(ship25qty1) <- 12  
  # 如果有空位，則用0填充
  ship25qty1[is.na(ship25qty1)] <- 0
  
  # 合併
  shipment1 <- rbind(ship23qty1,ship24qty1, ship25qty1)
  shipmentqty1 <- data.table(shipment1)
  
  # 找出長度
  len_wide_dt1 <- length( returnqty1[,-1])
  len_shipment1 <- length(shipmentqty1)
  
  # 如果 b 資料< a 資料，將 b 調整為與 a 長度相同
  if(len_shipment1 < len_wide_dt1) {
    shipmentqty2 <- rep(shipmentqty1[[ncol(shipmentqty1)]],each = len_wide_dt1-len_shipment1)  # 用 b 的最後一個值填充至 a 的長度
  }
  
  # 轉換成矩陣，每列是一個值
  result1 <- matrix(shipmentqty2, nrow = length(shipmentqty1[[length(shipmentqty1)]]), byrow = TRUE)
  
  shipmentqty1 <- cbind(shipmentqty1,result1)
  
  
  df_percent <- CRRdecimal
  df_percent[] <- lapply(CRRdecimal, function(x) {
    # 判斷是否為小數點數字，並忽略 NA
    if (any(x %% 1 != 0, na.rm = TRUE)) {
      # 將小數轉換為百分比格式
      return(paste0(formatC(x * 100, format = "f", digits = 2), "%"))
    } else {
      # 直接返回原值
      return(x)
    }
  })
  
  CRR <- cbind(returnqty[,1],df_percent)
  
  # 使用 gather() 函數將數據轉換為長格式
  data_long <- gather(CRR, key = "Lifetime", value = "SN", -Year)
  
  # 移除百分號並將其轉換為數字格式
  data_long$SN <- as.numeric(gsub("%", "", data_long$SN)) / 100
  
  # 確保 Lifetime 列按照數字順序排列
  data_long$Lifetime <- factor(data_long$Lifetime, 
                               levels = paste0("LT", 1:(length(CRR)-1)), 
                               ordered = TRUE)
  setDT(data_long)
  df_clean <- data_long %>%
    filter(!is.na(Lifetime) & !is.na(SN) & !is.infinite(Lifetime) & !is.infinite(SN))
  # 計算每個 Year 的最後一個有效的 SN 值
  last_valid_SN <- df_clean[!is.na(SN), .(last_SN = SN[.N]), by = Year]
  
  # 設置顯示條件：六的倍數和每年最後的有效值
  df_clean$show_label <- ifelse(
    (as.numeric(gsub("LT", "", df_clean$Lifetime)) %% 6 == 0) |
      (df_clean$SN == last_valid_SN$last_SN[match(df_clean$Year, last_valid_SN$Year)]),
    TRUE, 
    FALSE
  )
  #每月須改
  df_clean_filtered <- df_clean %>%
    # 篩選 2025SN 且 Lifetime 為 LT1
    filter(Year == "2025SN" & Lifetime == "LT5") %>%
    # 篩選 2024SN 且 Lifetime 為 LT1 到 LT13
    bind_rows(
      df_clean %>%
        filter(Year == "2024SN" & between(as.integer(substr(Lifetime, 3, 4)), 1, 17))
    ) %>%
    # 篩選 2023SN 且 Lifetime 為 LT1 到 LT25
    bind_rows(
      df_clean %>%
        filter(Year == "2023SN" & between(as.integer(substr(Lifetime, 3, 4)), 1, 29))
    )
  
  
  # 使用 ggplot2 繪製折線圖
  p <- ggplot(df_clean_filtered, aes(x = Lifetime, y = SN, color = Year, group = Year)) +
    geom_line(linewidth = 2.5, na.rm = TRUE) +  # 畫折線
    geom_point(size = 5, na.rm = TRUE) +  # 增加資料點大小
    geom_text(aes(label = ifelse(!is.na(SN) & show_label, paste0(round(SN * 100, 2), "%"), "")),
              vjust = ifelse(df_clean_filtered$Year == "2025SN", -3.5, ifelse(df_clean_filtered$Year == "2024SN", -1.2, -0.5)), hjust = 0.5, size = 9, show.legend = FALSE) +
    scale_color_manual(values = c("blue", "orange", "springgreen4")) +  # 更柔和的顏色搭配
    scale_y_continuous(labels = scales::percent_format(scale = 100)) +  # Y 軸顯示為百分比
    coord_cartesian(clip = "off") + 
    theme_minimal() +
    theme(
      axis.text.x = element_text(size = 16, hjust = 1, color = "black"),  # X 軸字體大小調整，無旋轉
      axis.text.y = element_text(size = 16, color = "black"),  # Y 軸字體大小調整，無粗體
      plot.margin = margin(30, 20, 20, 10),  # 圖表邊距
      legend.position = "bottom",  # 圖例位置
      legend.text = element_text(size = 14),  # 調整圖例文字大小
      legend.title = element_blank(),  # 移除圖例標題
      axis.title.y = element_blank(),  # 移除 Y 軸標題
      legend.background = element_rect(fill = "white", color = "black"),  # 設置圖例背景
      panel.grid.major = element_line(color = "gray", size = 0.5),  # 增加主網格線以提高可讀性
      panel.grid.minor = element_line(color = "gray", size = 0.2)  # 增加次網格線
    )
  
  #製作完整圖表
  result1 <- (CRRdecimal[1, ] - CRRdecimal[2, ]) / CRRdecimal[1, ]
  result2 <- (CRRdecimal[2, ] - CRRdecimal[3, ]) / CRRdecimal[2, ]
  # 3. 將小數轉回百分比
  result_percent1 <- paste0(round(result1 * 100, 2), "%")
  result_percent2 <- paste0(round(result2 * 100, 2), "%")
  result_percent1 <- c("24SN VS 25SN", result_percent1)
  result_percent2 <- c("25SN VS 24SN", result_percent2)
  result_percent1 <- as.data.table(result_percent1)
  result_percent2 <- as.data.table(result_percent2)
  result_percent1 <- transpose(result_percent1)
  result_percent2 <- transpose(result_percent2)
  
  returnqty1[,1] <- c("2023SN返修", "2024SN返修", "2025SN返修")
  shipmentqty1 <- cbind(c("2023SN出貨", "2024SN出貨", "2025SN出貨"),shipmentqty1)
  
  complete_table <- rbindlist(list(CRR, result_percent1, result_percent2, returnqty1, shipmentqty1[,1:13]), use.names = FALSE, fill = TRUE)
  # 將表格中的 "NA%" 替換成真正的 NA
  complete_table[complete_table == "NA%"] <- NA
  return(list(table = complete_table, plot = p))
}

SYSCRR <- data_table_and_plot(DataClean_SYS,SYSshipment23,SYSshipment24,SYSshipment25)
L6CRR <- data_table_and_plot(DataClean_L6,L6shipment23,L6shipment24,L6shipment25)
L10CRR <- data_table_and_plot(DataClean_L10,L10shipment23,L10shipment24,L10shipment25)
SBCRR <- data_table_and_plot(DataClean_SB,SBshipment23,SBshipment24,SBshipment25)

system <- rbind(SYSCRR$table[3,], L6CRR$table[3,], L10CRR$table[3,])
system[1, 1] <- "L6L10"
system[2, 1] <- "L6"
system[3, 1] <- "L10"

system[system == "NA%"] <- NA
# 需改
system <- system[,1:12]

df_clean <- system %>%
  mutate(across(-Year, ~as.numeric(gsub("%", "", .)) / 100))  # 除以 100 轉成小數

df_long <- pivot_longer(df_clean, cols = -Year, names_to = "LT", values_to = "Value")
setDT(df_long) 
last_valid_Value <- df_long[!is.na(Value), .(last_Value = Value[.N]), by = Year]

df_long$show_label <- ifelse(
  (as.numeric(gsub("LT", "", df_long$LT)) %% 6 == 0) | 
    (df_long$Value == last_valid_Value$last_Value[match(df_long$Year, last_valid_Value$Year)]),
  TRUE, 
  FALSE
)

# 需改
lt_order <- paste0("LT", 1:12)

# 將 LT 欄位轉為因子，並指定順序
df_long$LT <- factor(df_long$LT, levels = lt_order)

CRRcompare <- ggplot(df_long, aes(x = LT, y = Value, group = Year, color = Year)) +
  geom_line(linewidth = 2.5, na.rm = TRUE) +  # 畫折線
  geom_point(size = 5, na.rm = TRUE) +  # 增加資料點大小
  geom_text(aes(label = ifelse(show_label, paste0(round(Value * 100, 2), "%"), "")), 
            vjust = -1, size = 9, show.legend = FALSE) +
  scale_color_manual(values = c("blue", "orange", "springgreen4")) +  # 更柔和的顏色搭配
  scale_y_continuous(labels = scales::percent_format(scale = 100)) +  # Y 軸顯示為百分比
  coord_cartesian(clip = "off") + 
  theme_minimal() +
  ggtitle("2025SN System : WW CRR by Level") +
  theme(
    axis.text.x = element_text(size = 16, hjust = 1, color = "black"),  # X 軸字體大小調整，無旋轉
    axis.text.y = element_text(size = 16, color = "black"),  # Y 軸字體大小調整，無粗體
    plot.margin = margin(30, 20, 20, 10),  # 圖表邊距
    legend.position = "bottom",  # 圖例位置
    legend.text = element_text(size = 14),  # 調整圖例文字大小
    legend.title = element_blank(),  # 移除圖例標題
    axis.title.y = element_blank(),  # 移除 Y 軸標題
    legend.background = element_rect(fill = "white", color = "black"),  # 設置圖例背景
    panel.grid.major = element_line(color = "gray", size = 0.5),  # 增加主網格線以提高可讀性
    panel.grid.minor = element_line(color = "gray", size = 0.2),# 增加次網格線
    plot.title = element_text(size = 20, hjust = 0.5)
  ) 

ragg::agg_png("CRRcompare.png", width = 1600, height = 800)
print(CRRcompare)
dev.off()  # 關閉設備

# 讀取現有的 Excel 檔案
file_path <- "C:/Users/rex4_chen/Desktop/CRRtableplot.xlsx"
CRRtableplot <- loadWorkbook(file_path)
all_sheets <- getSheetNames(file_path)
history <- all_sheets[1]

# 刪除其他工作表
for (sheet in all_sheets[-1]) {
  removeWorksheet(CRRtableplot, sheet)
}
# 保存圖片為 PNG 格式
ragg::agg_png("L6L10CRR.png", width = 1600, height = 800)
print(SYSCRR$plot)
dev.off()  # 關閉設備
ragg::agg_png("L6CRR.png", width = 1600, height = 800)
print(L6CRR$plot)
dev.off()  # 關閉設備
ragg::agg_png("L10CRR.png", width = 1600, height = 800)
print(L10CRR$plot)
dev.off()  # 關閉設備
ragg::agg_png("SBCRR.png", width = 1600, height = 800)
print(SBCRR$plot)
dev.off()  # 關閉設備
# 讀取第一個工作表的資料，這個工作表不會改動
history <- read.xlsx(CRRtableplot, sheet = 1)

# 將新的資料寫入對應的工作表
#SYS
addWorksheet(CRRtableplot, "L6L10")
writeData(CRRtableplot, "L6L10", SYSCRR$table, startRow = 1, startCol = 1)
insertImage(CRRtableplot, "L6L10", "L6L10CRR.png", width = 6, height = 4, startRow = nrow(SYSCRR$table) + 3, startCol = 1)
#L6
addWorksheet(CRRtableplot, "L6")
writeData(CRRtableplot, "L6", L6CRR$table, startRow = 1, startCol = 1)
insertImage(CRRtableplot, "L6", "L6CRR.png", width = 6, height = 4, startRow = nrow(L6CRR$table) + 3, startCol = 1)
#L10
addWorksheet(CRRtableplot, "L10")
writeData(CRRtableplot, "L10", L10CRR$table, startRow = 1, startCol = 1)
insertImage(CRRtableplot, "L10", "L10CRR.png", width = 6, height = 4, startRow = nrow(L10CRR$table) + 3, startCol = 1)
#SB
addWorksheet(CRRtableplot, "SB")
writeData(CRRtableplot, "SB", SBCRR$table, startRow = 1, startCol = 1)
insertImage(CRRtableplot, "SB", "SBCRR.png", width = 6, height = 4, startRow = nrow(SBCRR$table) + 3, startCol = 1)
#CRRcompare
addWorksheet(CRRtableplot, "CRRcompare")
insertImage(CRRtableplot, "CRRcompare", "CRRcompare.png", width = 6, height = 4, startRow = nrow(SBCRR$table) + 3, startCol = 1)
CRRcompare
# 保存修改後的 Excel 檔案為新檔案
today_date <- format(Sys.Date(), "%Y%m%d")

file_path <- paste0("C:/Users/rex4_chen/Desktop/CRRtableplot", today_date, ".xlsx")
# 保存 Excel 文件
saveWorkbook(CRRtableplot, file_path, overwrite = TRUE)
