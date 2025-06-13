ShipmentData_ALL <- read_feather("C:/Users/rex4_chen/Desktop/feather/ShipmentData_ALL.feather")
ShipmentData_ALL$Month <- sprintf("%02d", as.numeric(ShipmentData_ALL$Month))  # 將月份轉為兩位數
ShipmentData_ALL$Year <- as.character(ShipmentData_ALL$Year)

# 新增 YM 列
ShipmentData_ALL$YM <- paste(ShipmentData_ALL$Year, ShipmentData_ALL$Month, sep = "/")
TTFR_ShipData <- subset(ShipmentData_ALL, YM >= "2019/11")
TTFR_ShipData <- TTFR_ShipData[Company != "ASUS", ]
TTFR_ShipData <- TTFR_ShipData[`BG...10` == "OP", ]
TTFR_ShipData[TTFR_ShipData$Country == "TÜRKIYE", "Country"] <- "TURKEY"

SVF <- TTFR_ShipData[`Product Line` %like% "SV|SF", ]
SVFCK <- TTFR_ShipData[`Product Line` %like% "SV|SF|SC|SK",]
SVF_ACC <- TTFR_ShipData[Country == "CHINA" & `Product Line` %like% "SV|SF",]
SB <- TTFR_ShipData[`Product Line` == "SB",]
SB_ACC <- TTFR_ShipData[Country == "CHINA" & `Product Line` == "SB", ]


summarize_and_pivot <- function(data) {
  # 分組並計算 Net QTY 加總
  result <- data %>%
    group_by(`BG...10`, Territory, Region, Branch, Country, YM) %>%
    summarise(`Net QTY` = sum(`Net QTY`, na.rm = TRUE), .groups = "drop") %>%
    pivot_wider(names_from = YM, values_from = `Net QTY`, values_fill = 0) %>%
    # 移除 BG...10, Region, Branch 欄位
    select(-`BG...10`, -Region, -Branch)
  
  date_cols <- setdiff(names(result), c("Territory", "Country")) # 非分組欄位
  sorted_cols <- c("Territory", "Country", sort(date_cols))     # 按日期排序
  
  result <- result[, sorted_cols]  # 按順序重新排列
  return(result)
}


SVFship <- summarize_and_pivot(SVF)
SVFCKship <- summarize_and_pivot(SVFCK)
SVF_ACCship <- summarize_and_pivot(SVF_ACC)
SBship <- summarize_and_pivot(SB)
SB_ACCship <- summarize_and_pivot(SB_ACC)

SVF_stock <- SVFship[, c("Territory", "Country")]
SVFCK_stock <- SVFCKship[, c("Territory", "Country")]
SVF_ACC_stock <- SVF_ACCship[, c("Territory", "Country")]
SB_stock <- SBship[, c("Territory", "Country")]
SB_ACC_stock <- SB_ACCship[, c("Territory", "Country")]

SVF_stockCT <- SVF_stock %>%
  mutate(Territory = if_else(Country == "TAIWAN", "TAIWAN",
                             if_else(Country == "CHINA", "CHINA", Territory)))
SVFCK_stockCT <- SVFCK_stock %>%
  mutate(Territory = if_else(Country == "TAIWAN", "TAIWAN",
                             if_else(Country == "CHINA", "CHINA", Territory)))
SVF_ACC_stockCT <- SVF_ACC_stock %>%
  mutate(Territory = if_else(Country == "TAIWAN", "TAIWAN",
                             if_else(Country == "CHINA", "CHINA", Territory)))
SB_stockCT <- SB_stock %>%
  mutate(Territory = if_else(Country == "TAIWAN", "TAIWAN",
                             if_else(Country == "CHINA", "CHINA", Territory)))
SB_ACC_stockCT <- SB_ACC_stock %>%
  mutate(Territory = if_else(Country == "TAIWAN", "TAIWAN",
                             if_else(Country == "CHINA", "CHINA", Territory)))

SVF_stockE <- SVF_stock %>%
  filter(Territory %in% c("EEMEA", "WE")) %>%  
  mutate(Territory = "EMEA")

SVFCK_stockE <- SVFCK_stock %>%
  filter(Territory %in% c("EEMEA", "WE")) %>%  
  mutate(Territory = "EMEA")

SVF_ACC_stockE <- SVF_ACC_stock %>%
  filter(Territory %in% c("EEMEA", "WE")) %>%  
  mutate(Territory = "EMEA")

SB_stockE <- SB_stock %>%
  filter(Territory %in% c("EEMEA", "WE")) %>%  
  mutate(Territory = "EMEA")

SB_ACC_stockE <- SB_ACC_stock %>%
  filter(Territory %in% c("EEMEA", "WE")) %>%  
  mutate(Territory = "EMEA")


STOCK <- data.frame(
  Territory = c("APAC", "TAIWAN", "BRAZIL", "CHINA", "EEMEA", "NA", "SA", "WE"),
  STOCK = c(3, 2, 3, 3, 5, 5, 5, 5)
)


SVF_stock <- SVF_stock %>%
  mutate(Territory = if_else(Territory == "China", "CHINA", 
                             if_else(Country == "TAIWAN", "TAIWAN", Territory))) %>%
  left_join(STOCK, by = "Territory")

SVFCK_stock <- SVFCK_stock %>%
  mutate(Territory = if_else(Territory == "China", "CHINA", 
                             if_else(Country == "TAIWAN", "TAIWAN", Territory))) %>%
  left_join(STOCK, by = "Territory")

SVF_ACC_stock <- SVF_ACC_stock %>%
  mutate(Territory = if_else(Territory == "China", "CHINA", 
                             if_else(Country == "TAIWAN", "TAIWAN", Territory))) %>%
  left_join(STOCK, by = "Territory")

SB_stock <- SB_stock %>%
  mutate(Territory = if_else(Territory == "China", "CHINA", 
                             if_else(Country == "TAIWAN", "TAIWAN", Territory))) %>%
  left_join(STOCK, by = "Territory")

SB_ACC_stock <- SB_ACC_stock %>%
  mutate(Territory = if_else(Territory == "China", "CHINA", 
                             if_else(Country == "TAIWAN", "TAIWAN", Territory))) %>%
  left_join(STOCK, by = "Territory")

SVF_mapping_data <- data.frame(
  GROUP = SVF_stock[[2]],  
  COUNTRY = SVF_stock[[2]], 
  check.names = FALSE,
  stringsAsFactors = FALSE  
)

SVFCK_mapping_data <- data.frame(
  GROUP = SVFCK_stock[[2]], 
  COUNTRY = SVFCK_stock[[2]],  
  check.names = FALSE,
  stringsAsFactors = FALSE  
)

SVF_ACC_mapping_data <- data.frame(
  GROUP = SB_ACC_stock[[2]],  
  COUNTRY = SB_ACC_stock[[2]], 
  check.names = FALSE,
  stringsAsFactors = FALSE  
)

SB_mapping_data <- data.frame(
  GROUP = SB_stock[[2]],  
  COUNTRY = SB_stock[[2]], 
  check.names = FALSE,
  stringsAsFactors = FALSE  
)

SB_ACC_mapping_data <- data.frame(
  GROUP = SB_ACC_stock[[2]],  
  COUNTRY = SB_ACC_stock[[2]],  
  check.names = FALSE,
  stringsAsFactors = FALSE  
)
SVF_stockCT <- SVF_stockCT %>%
  rename(GROUP = Territory, COUNTRY = Country)

SVFCK_stockCT <- SVFCK_stockCT %>%
  rename(GROUP = Territory, COUNTRY = Country)

SVF_ACC_stockCT <- SVF_ACC_stockCT %>%
  rename(GROUP = Territory, COUNTRY = Country)

SB_stockCT <- SB_stockCT %>%
  rename(GROUP = Territory, COUNTRY = Country)

SB_ACC_stockCT <- SB_ACC_stockCT %>%
  rename(GROUP = Territory, COUNTRY = Country)

SVF_mapping_data <- rbind(SVF_mapping_data, SVF_stockCT)
SVFCK_mapping_data <- rbind( SVFCK_mapping_data, SVFCK_stockCT)
SVF_ACC_mapping_data <- rbind(SVF_ACC_mapping_data, SVF_ACC_stockCT)
SB_mapping_data <- rbind(SB_mapping_data, SB_stockCT)
SB_ACC_mapping_data <- rbind(SB_ACC_mapping_data, SB_ACC_stockCT)

SVF_stockE <- SVF_stockE %>%
  rename(GROUP = Territory, COUNTRY = Country)

SVFCK_stockE <- SVFCK_stockE %>%
  rename(GROUP = Territory, COUNTRY = Country)

SVF_ACC_stockE <- SVF_ACC_stockE %>%
  rename(GROUP = Territory, COUNTRY = Country)

SB_stockE <- SB_stockE %>%
  rename(GROUP = Territory, COUNTRY = Country)

SB_ACC_stockE <- SB_ACC_stockE %>%
  rename(GROUP = Territory, COUNTRY = Country)

SVF_mapping_data <- rbind(SVF_mapping_data, SVF_stockE)
SVFCK_mapping_data <- rbind( SVFCK_mapping_data, SVFCK_stockE)
SVF_ACC_mapping_data <- rbind(SVF_ACC_mapping_data, SVF_ACC_stockE)
SB_mapping_data <- rbind(SB_mapping_data, SB_stockE)
SB_ACC_mapping_data <- rbind(SB_ACC_mapping_data, SB_ACC_stockE)


new_SVF <- data.frame(
  GROUP = "WW",                
  COUNTRY = SVF_stock[[2]],     
  stringsAsFactors = FALSE      
)

SVF_mapping_data <- rbind(SVF_mapping_data, new_SVF)

new_SVFCK <- data.frame(
  GROUP = "WW",                
  COUNTRY = SVFCK_stock[[2]],     
  stringsAsFactors = FALSE      
)

SVFCK_mapping_data <- rbind(SVFCK_mapping_data, new_SVFCK)

new_SVF_ACC <- data.frame(
  GROUP = "WW",                
  COUNTRY = SVF_ACC_stock[[2]],     
  stringsAsFactors = FALSE      
)

SVF_ACC_mapping_data <- rbind(SVF_ACC_mapping_data, new_SVF_ACC)

new_SB <- data.frame(
  GROUP = "WW",                
  COUNTRY = SB_stock[[2]],     
  stringsAsFactors = FALSE      
)

SB_mapping_data <- rbind(SB_mapping_data, new_SB)

new_SB_ACC <- data.frame(
  GROUP = "WW",                
  COUNTRY = SB_ACC_stock[[2]],     
  stringsAsFactors = FALSE      
)

SB_ACC_mapping_data <- rbind(SB_ACC_mapping_data, new_SB_ACC)

SVFship <- SVFship %>%
  left_join(SVF_stock, by = "Country") %>%  # 根據 Country 匹配 STOCK
  mutate(
    WARRANTY = 3                             # 新增 WARRANTY 欄位，填滿 3
  ) %>%
  select(Country, STOCK, WARRANTY, everything(), -Territory.x, -Territory.y)

SVFCKship <- SVFCKship %>%
  left_join(SVFCK_stock, by = "Country") %>%  # 根據 Country 匹配 STOCK
  mutate(
    WARRANTY = 3                             # 新增 WARRANTY 欄位，填滿 3
  ) %>%
  select(Country, STOCK, WARRANTY, everything(), -Territory.x, -Territory.y)

SVF_ACCship <- SVF_ACCship %>%
  left_join(SVF_ACC_stock, by = "Country") %>%  # 根據 Country 匹配 STOCK
  mutate(
    WARRANTY = 3                             # 新增 WARRANTY 欄位，填滿 3
  ) %>%
  select(Country, STOCK, WARRANTY, everything(), -Territory.x, -Territory.y)

SBship <- SBship %>%
  left_join(SB_stock, by = "Country") %>%  # 根據 Country 匹配 STOCK
  mutate(
    WARRANTY = 3                             # 新增 WARRANTY 欄位，填滿 3
  ) %>%
  select(Country, STOCK, WARRANTY, everything(), -Territory.x, -Territory.y)

SB_ACCship <- SB_ACCship %>%
  left_join(SB_ACC_stock, by = "Country") %>%  # 根據 Country 匹配 STOCK
  mutate(
    WARRANTY = 3                             # 新增 WARRANTY 欄位，填滿 3
  ) %>%
  select(Country, STOCK, WARRANTY, everything(), -Territory.x, -Territory.y)

colnames(SVFship)[colnames(SVFship) == "Country"] <- "COUNTRY"
colnames(SVFCKship)[colnames(SVFCKship) == "Country"] <- "COUNTRY"
colnames(SVF_ACCship)[colnames(SVF_ACCship) == "Country"] <- "COUNTRY"
colnames(SBship)[colnames(SBship) == "Country"] <- "COUNTRY"
colnames(SB_ACCship)[colnames(SB_ACCship) == "Country"] <- "COUNTRY"

# 創建 Excel 工作簿
SVF_Shipment <- createWorkbook()

# 新增第一個工作表 (Mapping)
addWorksheet(SVF_Shipment, "Mapping")
writeData(SVF_Shipment, "Mapping", unique(SVF_mapping_data))

# 新增第二個工作表 (ShipQty)
addWorksheet(SVF_Shipment, "ShipQty")
writeData(SVF_Shipment, "ShipQty", SVFship)

# 儲存 Excel 檔案
saveWorkbook(SVF_Shipment, "C:/Users/rex4_chen/Desktop/SVF_Shipment_202501.xlsx", overwrite = TRUE)


# 創建 Excel 工作簿
SVFCK_Shipment <- createWorkbook()

# 新增第一個工作表 (Mapping)
addWorksheet(SVFCK_Shipment, "Mapping")
writeData(SVFCK_Shipment, "Mapping", unique(SVFCK_mapping_data))

# 新增第二個工作表 (ShipQty)
addWorksheet(SVFCK_Shipment, "ShipQty")
writeData(SVFCK_Shipment, "ShipQty", SVFCKship)

# 儲存 Excel 檔案
saveWorkbook(SVFCK_Shipment, "C:/Users/rex4_chen/Desktop/SVFCK_Shipment_202501.xlsx", overwrite = TRUE)


# 創建 Excel 工作簿
SVF_ACC_Shipment <- createWorkbook()

# 新增第一個工作表 (Mapping)
addWorksheet(SVF_ACC_Shipment, "Mapping")
writeData(SVF_ACC_Shipment, "Mapping", unique(SVF_ACC_mapping_data))

# 新增第二個工作表 (ShipQty)
addWorksheet(SVF_ACC_Shipment, "ShipQty")
writeData(SVF_ACC_Shipment, "ShipQty", SVF_ACCship)

# 儲存 Excel 檔案
saveWorkbook(SVF_ACC_Shipment, "C:/Users/rex4_chen/Desktop/SVF_ACC_Shipment_202501.xlsx", overwrite = TRUE)


# 創建 Excel 工作簿
SB_Shipment <- createWorkbook()

# 新增第一個工作表 (Mapping)
addWorksheet(SB_Shipment, "Mapping")
writeData(SB_Shipment, "Mapping", unique(SB_mapping_data))

# 新增第二個工作表 (ShipQty)
addWorksheet(SB_Shipment, "ShipQty")
writeData(SB_Shipment, "ShipQty", SBship)

# 儲存 Excel 檔案
saveWorkbook(SB_Shipment, "C:/Users/rex4_chen/Desktop/SB_Shipment_20250421.xlsx", overwrite = TRUE)

# 創建 Excel 工作簿
SB_ACC_Shipment <- createWorkbook()

# 新增第一個工作表 (Mapping)
addWorksheet(SB_ACC_Shipment, "Mapping")
writeData(SB_ACC_Shipment, "Mapping", unique(SB_ACC_mapping_data))

# 新增第二個工作表 (ShipQty)
addWorksheet(SB_ACC_Shipment, "ShipQty")
writeData(SB_ACC_Shipment, "ShipQty", SB_ACCship)

# 儲存 Excel 檔案
saveWorkbook(SB_ACC_Shipment, "C:/Users/rex4_chen/Desktop/SB_ACC_Shipment_20250421.xlsx", overwrite = TRUE)

