library(openxlsx)
library(writexl)
library(readxl)
library(arrow)
library(data.table)
library(dplyr)
library(tidyr)
library(shiny)
library(scales)

# 分子
DataClean_raw.dt_d <- read_feather("C:/Users/rex4_chen/Desktop/feather/DataClean_raw.dt_d.feather")

repair_data <- DataClean_raw.dt_d[, c("PRODUCT_LINE", "MUC_MODULE", "MODEL_NAME", "ACTION_CODE", 
                                      "Problem_Combine", "ORG_REPLACED_PART_DESC", 
                                      "Mix_Repair_LMD", "ORG_MODEL_PART_NO", "REPAIR_LEVEL", "Close_Y", "Category")]
# 分母
ShipmentData_ALL <- read_feather("C:/Users/rex4_chen/Desktop/feather/ShipmentData_ALL.feather")
ori_raw.dt <- data.table(ShipmentData_ALL)
raw.dt_o <- ori_raw.dt[!duplicated(ori_raw.dt),]

raw.dt_o <- raw.dt_o[BU == "Server",]
raw.dt_PRO <- raw.dt_o[!(grepl("E500|E800|E900|^EG", Model))]
WS_Mapping <-read_excel("D:/PQA/Raw_Data/Mapping/Workstation_list.xlsx")
setDT(WS_Mapping)
raw.dt_PRO <- raw.dt_PRO[`Model` %in% WS_Mapping$Model, WS := "Y"]
raw.dt_PRO <- raw.dt_PRO[is.na(WS), 
                         WS := "N"]
raw.dt_PRO <- raw.dt_PRO[WS == "N",]

raw.dt_PRO[, Month := as.character(Month)]
raw.dt_PRO[, Year := as.character(Year)]

raw.dt_PRO[`Product Line`%in% c("S","SV","SF") & `Product Desc`%like% c("WOC"), L6_L10 := "準系統"]
Server_Level <- read_excel("D:/PQA/Raw_Data/Mapping/Mapping_Table_SYS_Level.xlsx", sheet = "SYS_Level")
raw.dt_PRO[,L6_L10 := Server_Level$SYS_Level[match(raw.dt_PRO$'Part Number',Server_Level$Part_Number)]]

raw.dt_PRO[, Month := as.character(Month)]
raw.dt_PRO[, Year := as.character(Year)]
raw.dt_PRO$model_new <- gsub("(.*?)(/.*)?", "\\1", raw.dt_PRO$`Product Desc`)

shipment_data <- raw.dt_PRO[, c("Product Line", "model_new", "L6_L10", "Net QTY", "Year")]
colnames(shipment_data) <- c("Product_Line", "Model", "L6_L10", "Net_QTY", "Year")

ui <- fluidPage(
  # 美化整體頁面
  theme = bslib::bs_theme(
    bg = "#f7f7f7", # 背景色
    fg = "#333333", # 文字顏色
    primary = "#0066cc", # 主題顏色
    secondary = "#ff6600" # 副主題顏色
  ),
  
  # 標題
  titlePanel(
    title = div(
      style = "font-weight: bold; color: #0066cc;", 
      "返修與出貨篩選工具"
    )
  ),
  
  fluidRow(
    # 左邊篩選返修數據
    column(
      width = 6,
      h3("篩選分子（返修數據）", style = "color: #0066cc;"),
      wellPanel(
        style = "background-color: #ffffff; border-radius: 10px; padding: 20px;",
        selectizeInput("PRODUCT_LINE", "PRODUCT_LINE:", 
                       choices = c("ALL", "", sort(unique(as.character(repair_data$PRODUCT_LINE)))), 
                       selected = "ALL", 
                       multiple = FALSE),
        selectizeInput("module", "MUC_MODULE:", 
                       choices = c("ALL", "", sort(unique(as.character(repair_data$MUC_MODULE)))), 
                       selected = "ALL", 
                       multiple = FALSE),
        selectizeInput("model", "MODEL_NAME:", 
                       choices = c("", sort(unique(as.character(repair_data$MODEL_NAME)))), 
                       selected = NULL, 
                       multiple = TRUE),
        selectizeInput("problem_combine", "Problem_Combine:", 
                       choices = c("", sort(unique(as.character(repair_data$Problem_Combine)))), 
                       selected = NULL, 
                       multiple = FALSE),
        selectizeInput("replaced_part_desc", "ORG_REPLACED_PART_DESC:", 
                       choices = c("ALL", "", sort(unique(as.character(repair_data$ORG_REPLACED_PART_DESC)))), 
                       selected = "ALL", 
                       multiple = FALSE),
        selectizeInput("part_no", "ORG_MODEL_PART_NO:", 
                       choices = c("", sort(unique(as.character(repair_data$ORG_MODEL_PART_NO)))), 
                       selected = NULL, 
                       multiple = FALSE),
        selectizeInput("category", "Category:", 
                       choices = c("ALL", "", sort(unique(as.character(repair_data$Category)))), 
                       selected = "ALL", 
                       multiple = FALSE),
        radioButtons("repair_type", "修復類型:",
                     choices = c("全部" = "all", "換料" = "R10", "非換料" = "non-R10"),
                     selected = "all"),
        sliderInput("close_y_range", "選擇年份範圍(Close_Y):",
                    min = 2020, max = 2024,
                    value = c(2020, 2024), step = 1,
                    sep = ""),
        checkboxInput("exclude_multiple_repairs", "排除一筆多修", value = TRUE),
        checkboxInput("exclude_cid_oow", "排除 REPAIR_LEVEL 為 'CID/OOW'", value = TRUE),
        actionButton("filter_repair", "篩選返修數據", class = "btn-primary")
      )
    ),
    
    # 右邊篩選出貨數據
    column(
      width = 6,
      h3("篩選分母（出貨數據）", style = "color: #0066cc;"),
      wellPanel(
        style = "background-color: #ffffff; border-radius: 10px; padding: 20px;",
        selectizeInput("product_line", "Product Line:", 
                       choices = c("", sort(unique(as.character(shipment_data$Product_Line)))), 
                       selected = NULL, 
                       multiple = TRUE),
        selectizeInput("model_ship", "Model:", 
                       choices = c("", sort(unique(as.character(shipment_data$Model)))), 
                       selected = NULL, 
                       multiple = TRUE),
        selectizeInput("l6_l10", "L6/L10:", 
                       choices = c("", sort(unique(as.character(shipment_data$L6_L10)))), 
                       selected = NULL, 
                       multiple = FALSE),
        sliderInput("year_range", "選擇年份範圍:", 
                    min = 2020, max = 2024, 
                    value = c(2020, 2024), step = 1, 
                    sep = ""),
        actionButton("filter_shipment", "篩選出貨數據", class = "btn-primary")
      ),
      
      mainPanel(
        style = "background-color: #f0f0f0; padding: 20px; border-radius: 10px;",
        textOutput("repair_count"),  # 顯示返修數量
        textOutput("shipment_count"),  # 顯示出貨數量
        textOutput("fail_rate"), # 顯示fail rate
        h4("返修分組數量", style = "color: #0066cc;"),
        tableOutput("repair_group_count"),
        downloadButton("download_repair_table", "匯出 Excel", class = "btn-secondary")
      )
    )
  ),
  
  # 加入頁面底部的額外設計
  footer = div(
    style = "text-align: center; padding: 20px; background-color: #0066cc; color: white; font-size: 14px;",
    "Data Engineering & Analytics Tools | Powered by Shiny"
  )
)

server <- function(input, output) {
  
  # 篩選返修數據
  filtered_repair <- eventReactive(input$filter_repair, {
    
    filtered <- repair_data
    
    if (input$exclude_cid_oow) {
      filtered <- filtered[is.na(REPAIR_LEVEL) | REPAIR_LEVEL != "CID/OOW", ]
    }
    if (input$repair_type == "R10") {
      filtered <- filtered[grepl("R10", ACTION_CODE, ignore.case = TRUE), ]
    } else if (input$repair_type == "non-R10") {
      filtered <- filtered[!grepl("R10", ACTION_CODE, ignore.case = TRUE), ]
    }
    if (input$module != "ALL" && input$module != "") {
      filtered <- filtered[MUC_MODULE == input$module, ]
    }
    if (input$PRODUCT_LINE != "ALL" && input$PRODUCT_LINE != "") {
      filtered <- filtered[PRODUCT_LINE == input$PRODUCT_LINE, ]
    }
    if (length(input$model) > 0) {  # 確保有選擇 Model Name
      filtered <- filtered[MODEL_NAME %in% input$model, ]
    }
    if (input$problem_combine != "") {
      filtered <- filtered[grepl(input$problem_combine, Problem_Combine, ignore.case = TRUE), ]
    }
    if (input$replaced_part_desc != "ALL" && input$replaced_part_desc != "") {
      filtered <- filtered[grepl(input$replaced_part_desc, ORG_REPLACED_PART_DESC, ignore.case = TRUE), ]
    }
    if (input$part_no != "") {
      filtered <- filtered[ORG_MODEL_PART_NO == input$part_no, ]
    }
    if (input$category != "ALL" && input$category != "") {
      filtered <- filtered[grepl(input$category, Category, ignore.case = TRUE), ]
    }
    if (input$exclude_multiple_repairs) {
      filtered <- filtered[Mix_Repair_LMD == "Main_Repair", ]
    }
    year_min <- input$close_y_range[1]
    year_max <- input$close_y_range[2]
    filtered <- filtered[Close_Y >= year_min & Close_Y <= year_max, ]
    filtered
  })
  
  # 篩選出貨數據
  filtered_shipment <- eventReactive(input$filter_shipment, {
    filtered <- shipment_data
    
    if (length(input$product_line) > 0) {
      filtered <- filtered[Product_Line %in% input$product_line, ]
    }
    if (length(input$model_ship) > 0) {
      filtered <- filtered[Model %in% input$model_ship, ]
    }
    if (input$l6_l10 != "") {
      filtered <- filtered[grepl(input$l6_l10, L6_L10, ignore.case = TRUE), ]
    }
    year_min <- input$year_range[1]
    year_max <- input$year_range[2]
    filtered <- filtered[Year >= year_min & Year <= year_max, ]
    filtered
  })
  
  # 計算並顯示分組數據表
  output$repair_group_count <- renderTable({
    data <- filtered_repair()
    shipment_count <- sum(filtered_shipment()$Net_QTY, na.rm = TRUE)
    
    if (shipment_count == 0) {
      return(data.frame(提示 = "分母數量為 0，無法計算 Fail Rate"))
    }
    if (input$module != "ALL" && input$module != "") {
      group_by_column <- "MUC_MODULE"
    } else if (input$category != "ALL" && input$category != "") {
      group_by_column <- "Category"
    } else if (input$replaced_part_desc != "ALL" && input$replaced_part_desc != "") {
      group_by_column <- "ORG_REPLACED_PART_DESC"
    } else {
      return(data.frame(提示 = "未選擇有效的分組條件"))
    }
    
    grouped_data <- data %>% 
      group_by(.data[[group_by_column]]) %>% 
      summarise(
        Count = n(),
        Fail_Rate = percent(Count / shipment_count, accuracy = 0.01),
        .groups = "drop"
      ) %>%
      arrange(desc(Count))
    
    grouped_data
  })
  
  # 顯示分子數量
  output$repair_count <- renderText({
    repair_count <- nrow(filtered_repair())
    paste("分子數字（返修數量）: ", repair_count)
  })
  
  # 顯示分母數量
  output$shipment_count <- renderText({
    shipment_count <- sum(filtered_shipment()$Net_QTY, na.rm = TRUE)
    paste("分母數字（出貨數量）: ", shipment_count)
  })
  
  # 計算並顯示 Fail Rate
  output$fail_rate <- renderText({
    repair_count <- nrow(filtered_repair())
    shipment_count <- sum(filtered_shipment()$Net_QTY, na.rm = TRUE)
    if (shipment_count == 0) {
      fail_rate <- NA
    } else {
      fail_rate <- (repair_count / shipment_count) * 100
    }
    paste("Fail Rate: ", round(fail_rate, 2), "%")
  })

  
  # 匯出分組數據表
  output$download_repair_table <- downloadHandler(
    filename = function() {
      paste0("repair_group_count_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      data <- filtered_repair()
      shipment_count <- sum(filtered_shipment()$Net_QTY, na.rm = TRUE)
      
      if (shipment_count == 0) {
        
        grouped_data <- data.frame(提示 = "分母數量為 0，無法計算 Fail Rate")
      } else if (input$module != "ALL" && input$module != "") {
        group_by_column <- "MUC_MODULE"
      } else if (input$category != "ALL" && input$category != ""){
        group_by_column <- "Category"
      } else if (input$replaced_part_desc != "ALL" && input$replaced_part_desc != "") {
        group_by_column <- "ORG_REPLACED_PART_DESC"
      } else {
        grouped_data <- data.frame(提示 = "未選擇有效的分組條件")
      }
      
      if (exists("group_by_column")) {
        grouped_data <- data %>%
          group_by(.data[[group_by_column]]) %>%
          summarise(
            Count = n(),
            Fail_Rate = percent(Count / shipment_count, accuracy = 0.01),
            .groups = "drop"
          ) %>%
          arrange(desc(Count))
      }
      
      write.xlsx(grouped_data, file)
    }
  )
}

# 啟動 Shiny 應用程式
shinyApp(ui = ui, server = server)
