library(shiny)
library(openxlsx)
library(dplyr)
library(shinyFeedback)

#setup uploaded file size limit
options(shiny.maxRequestSize = 1 * 1024^2) # 1 * 1024^2 = 1 Mb)

ui <- fluidPage(
  titlePanel("Compare Two Excel Spreadsheets"),
  
  useShinyFeedback(),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("oldFile", "Upload Old Excel File", accept = c(".xlsx", ".xls")),
      fileInput("newFile", "Upload New Excel File", accept = c(".xlsx", ".xls")),
      uiOutput("keySelectUI"),
      uiOutput("actionbtn"),
      htmlOutput("errorMsg")
    ),
    
    mainPanel(
      tableOutput("resultTable"),
      hr(),
      downloadButton("downloadData", "Download Merged Data")
    )
  )
)

server <- function(input, output, session) {
  #check file ext function
  checkext <- function(x) {
    substr(x , nchar(x) - 4, nchar(x)) == ".xlsx"
  }

  oldData <- reactive({
    feedbackWarning("oldFile", !checkext(input$oldFile$name), "file is not xlsx")
    req(checkext(input$oldFile$name))
    read.xlsx(input$oldFile$datapath)
  })
  
  newData <- reactive({
    feedbackWarning("newFile", !checkext(input$newFile$name), "file is not xlsx")
    req(checkext(input$newFile$name))
    read.xlsx(input$newFile$datapath)
  })

  observe({
    if (!identical(names(oldData()),names(newData()))) {
      output$errorMsg <- renderText("column names are diff")
      output$actionbtn <- NULL
      output$keySelectUI <- NULL
    }
    else{
      output$actionbtn <- renderUI(actionButton("action", "Compare"))
      output$keySelectUI <- renderUI(selectInput("key", "Select Key Column for Merging", choices = names(oldData()), multiple = TRUE))
      output$errorMsg <- NULL
    }
  })
  
  # Function to compare row values
  compare_row <- function(row_old, row_new, common_cols) {
    differences <- c()
    #26Mar25 added code
    if (all(sapply(row_old, is.na)) == TRUE) {
      differences = "data only in new file"
    } else if (all(sapply(row_new, is.na)) == TRUE) {
      differences = "data only in old file"
    } else {
      for (col in common_cols) {
        val_old <- row_old[[col]]
        val_new <- row_new[[col]]
        
        # Handle NA values separately
        if (is.na(val_old) && is.na(val_new)) next
        if ((is.na(val_old) && !is.na(val_new)) || (!is.na(val_old) && is.na(val_new)) || as.character(val_old) != as.character(val_new)) {
          differences <- c(differences, paste0(col, ": old=", val_old, ", new=", val_new))
        }
      }
      if (length(differences) == 0) return(NA_character_)
      paste(differences, collapse = "; ")
    }
  }
  
  # Perform comparison and merging
  observeEvent(input$action, {
    checkvaruniq <- oldData() %>%
      select(input$key)

    if (nrow(oldData()) != nrow(unique(checkvaruniq))) {
      #output$errorMsg <- renderText("selected variable are not unique")
      output$errorMsg <- renderText(
        HTML("<span style='color: red; font-size: 20px;'>Selected variables are not unique </span>")
      )
    }
    else{
      merged <- merge(oldData(), newData(), by = input$key, all = TRUE, suffixes = c(".old", ".new"))
      
      # Find common columns (excluding the key column)
      common_cols <- setdiff(names(oldData()), input$key)
      
      # Create "Differences" column
      differences <- apply(merged, 1, function(row) {
        row_old <- row[paste0(common_cols, ".old")]
        row_new <- row[paste0(common_cols, ".new")]
        
        # Convert to named list for easier handling
        #row_old_list <- as.list(row_old) #25Mar25 change to comment, below line updated
        names(row_old) <- common_cols
        #row_new_list <- as.list(row_new) #25Mar25 change to comment, below line updated
        names(row_new) <- common_cols
        
        compare_row(row_old, row_new, common_cols)
      })
      
      merged$Differences <- differences
      output$resultTable <- renderTable(merged)
    }

  })
  
  # Download merged result
  output$downloadData <- downloadHandler(
    filename = function() {
      paste("merged_comparison_", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      req(mergedData())
      write.csv(mergedData(), file, row.names = FALSE)
    }
  )
}

shinyApp(ui = ui, server = server)
