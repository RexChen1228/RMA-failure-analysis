library(openxlsx);library(data.table);library(stringr);library(ISOweek);library(zoo)
library(arrow)
##import data with RDS
#update original QC raw data convert to RDS (DL RDS)
##import data with RDS 
#yearly data
DataClean_raw.dt_d <- read_feather("C:/Users/rex4_chen/Desktop/feather/DataClean_raw.dt_d.feather")

ori_data <- data.table(DataClean_raw.dt_d)

#2019/07 - 2020/06(Warranty 2-3yrs)
#SN_2_month<-c("K7","K8","K9","KA","KB","KC","L1","L2","L3","L4","L5","L6")
#2024/2 - 2025/03(Warranty 2-3yrs)
SN_2_month<-c("M2","M3","M4","M5","M6","M7","M8","M9","MA","MB","MC", "N1")


#2019/07 - 2023/08(Warranty 2-3yrs),
# close_month <- c(sprintf("%d/%02.f", rep(2019     ,each= 6), 7:12), # 6 months
#                  sprintf("%d/%02.f", rep(2020:2022,each= 12), 1:12),# 36 months
#                  sprintf("%d/%02.f", rep(2023     ,each= 8), 1:8)) # 8 months

#2021/02 - 2025/03(Warranty 2-3yrs),
close_month <- c(sprintf("%d/%02.f", rep(2021     ,each= 11), 2:12), # 11 months
                 sprintf("%d/%02.f", rep(2022:2024,each= 12), 1:12),#  36months
                 sprintf("%d/%02.f", rep(2025     ,each= 3), 1:3)) # 3 months

ori_data <- ori_data[, ':=' (Close_M = format(as.Date(CLOSED_DATE), "%Y/%m"), 
                             SN_2 = substr(ORG_SN, 1, 2))]

# PRODUCT_LINE : SB(SW)/SV(SF)/SC/SK (Shipment during 2017/07~2018/06 by product line)
# Server all

data <- ori_data

# #2019/07 - 2020/06
# F6_EMS_SHIP<-c(27745,27714, 26602,28169,33800,27427,25606,20551,25355,33032,40652,37650)

#2021/02 - 2022/01
F6_EMS_SHIP<-c(22308, 27249, 26513, 43845, 30486, 35894, 33992, 48429, 45975, 52515, 44283, 38888)


# # Server all(exclude SW)
# data <- ori_data
# data <- data[!PRODUCT_LINE %in% "SW", ]
# data <- data[!c(PRODUCT_LINE %in% c("2016","2017") & PRODUCT_LINE %in% "SB" & ORG_MODEL_PART_DESC %like% "WS"), ]
# F6_EMS_SHIP<-c(29035,32701,33701,28607,32127,28079,33249,47244,28447,32386,25026,32984)


##CHECK
table(data$PRODUCT_LINE)
table(data$LMD_PART_GROUP)

LMD_part_name<-names(sort(table(data[LMD_PART_GROUP != ""& SN_2 %in% SN_2_month & Close_M %in% close_month,LMD_PART_GROUP]),decreasing=T))
pattern<-matrix(,39,length(LMD_part_name))
colnames(pattern)<-LMD_part_name
General_PCFR<-c()

wb <- createWorkbook()

header_style <- createStyle(fontSize = 12, textDecoration = "bold", halign = "center", valign = "center")

for(i in 1:length(LMD_part_name)){
  addWorksheet(wb, LMD_part_name[i])
  #i<-1
  print(LMD_part_name[i])
  run_data<-data[LMD_PART_GROUP == LMD_part_name[i] & SN_2 %in% SN_2_month & Close_M %in% close_month ,]
  original_m<-table(factor(run_data[,Close_M],levels=close_month),factor(run_data[,SN_2],levels=SN_2_month))
  
  #算39個  matrix(,39,length(SN_2_month))
  lifetime_m<-matrix(,39,length(SN_2_month))
  colnames(lifetime_m)<-SN_2_month
  index<-1:100
  
  #算39個 index[j+38]
  for(j in 1:length(SN_2_month)){
    lifetime_m[,j]<-original_m[index[j]:index[j+38],j] 
  }
     
  lifetime_by_month<-rowSums(lifetime_m)
  General_PCFR[i]<-sum(lifetime_by_month)/sum(F6_EMS_SHIP)
  lifetime_by_cum<-cumsum(lifetime_by_month)
  pattern[,i]<-lifetime_by_cum/tail(lifetime_by_cum,1)
  
  EMS_output<-t(F6_EMS_SHIP)
  colnames(EMS_output)<-SN_2_month
  
  if (LMD_part_name[i] == "02*BGA") {
    # 新增 "02_BGA" 工作表
    addWorksheet(wb, "02_BGA")
    # 寫入 EMS_output，從第 2 欄開始
    writeData(wb, "02_BGA", EMS_output, startCol = 2, header = TRUE, headerStyle = header_style)
    # 寫入 original_m，從第 3 行開始
    writeData(wb, "02_BGA", as.data.frame.matrix(original_m), startRow = 3, rowNames = TRUE, header = TRUE, headerStyle = header_style)
    # 寫入 lifetime_m，從第 56 行開始
    writeData(wb, "02_BGA", as.data.frame.matrix(lifetime_m), startRow = 56, rowNames = TRUE, header = TRUE, headerStyle = header_style)
  }else if(LMD_part_name[i]=="06*BGA"){
    addWorksheet(wb, "06_BGA")
    # 寫入 EMS_output，從第 2 欄開始
    writeData(wb, "06_BGA", EMS_output, startCol = 2, header = TRUE, headerStyle = header_style)
    # 寫入 original_m，從第 3 行開始
    writeData(wb, "06_BGA", as.data.frame.matrix(original_m), startRow = 3, rowNames = TRUE, header = TRUE, headerStyle = header_style)
    # 寫入 lifetime_m，從第 56 行開始
    writeData(wb, "06_BGA", as.data.frame.matrix(lifetime_m), startRow = 56, rowNames = TRUE, header = TRUE, headerStyle = header_style)
  }else{
    writeData(wb, LMD_part_name[i], EMS_output, startCol = 2)
    writeData(wb, LMD_part_name[i], as.data.frame.matrix(original_m), startRow = 3, colNames = TRUE, rowNames = TRUE)
    writeData(wb, LMD_part_name[i], "original_matrix", startRow = 3, startCol = 1)
    writeData(wb, LMD_part_name[i], as.data.frame.matrix(lifetime_m) , startRow = 56, colNames = TRUE, rowNames = TRUE)
    writeData(wb, LMD_part_name[i], "lifefime_matrix", startRow = 56, startCol = 1)
  }
  
  
  
}

names(General_PCFR)<-LMD_part_name
output_General_PCFR<-t(General_PCFR )
rownames(output_General_PCFR)<-"General_PCFR"
addWorksheet(wb, "pattern & General PCFR")
writeData(wb, "pattern & General PCFR", output_General_PCFR, startRow = 1, colNames = TRUE, rowNames = TRUE)
writeData(wb, "pattern & General PCFR", pattern, startCol = 1, startRow = 3, colNames = TRUE, rowNames = TRUE)
todaydate <- format(Sys.Date(), "%Y%m%d")
file_path <-  paste0("C:/Users/rex4_chen/Desktop/PCMR_", todaydate,".xlsx")

saveWorkbook(wb, file_path, overwrite = TRUE)
 