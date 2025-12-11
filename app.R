library(shiny)
library(openxlsx)
library(dplyr)
library(shinyFeedback)

#setup uploaded file size limit
options(shiny.maxRequestSize = 4 * 1024^2) # 1 * 1024^2 = 1 Mb)

ui <- fluidPage(
  titlePanel("Compare Two Excel Spreadsheets"),
  
  useShinyFeedback(),
  
  sidebarLayout(
    sidebarPanel(
      fileInput("oldFile", "Upload Original Excel File", accept = c(".xlsx", ".xls")),
      fileInput("newFile", "Upload Updated Excel File to Compare", accept = c(".xlsx", ".xls")),
      uiOutput("keySelectUI"),
      uiOutput("actionbtn"),
      br(),
      htmlOutput("errorMsg"),
      uiOutput("downloadbtn"),
    ),
    
    mainPanel(
      tableOutput("resultTable"),
      hr(),
    )
  )
)

server <- function(input, output, session) {
  #check file ext function
  checkext <- function(x) {
    substr(x , nchar(x) - 4, nchar(x)) == ".xlsx"
  }
  
  errormsg_f <- function(y) {
    output$errorMsg <- renderText(HTML("<span style='color: red; font-size: 15px;'>",y,"</span>"))
    output$actionbtn <- NULL
    output$keySelectUI <- NULL
  }

  oldData <- reactive({
    feedbackDanger("oldFile", !checkext(input$oldFile$name), "ONLY xlsx format file is accepted")
    output$errorMsg <- NULL
    req(checkext(input$oldFile$name))
    read.xlsx(input$oldFile$datapath)
  })
  
  newData <- reactive({
    feedbackDanger("newFile", !checkext(input$newFile$name), "ONLY xlsx format file is accepted")
    output$errorMsg <- NULL
    req(checkext(input$newFile$name))
    read.xlsx(input$newFile$datapath)
  })

  observe({
    if (!identical(sort(names(oldData())),sort(names(newData())))) {
      errormsg_f("Error: The two files do not have the same columns names.")
    }
    else if (length(names(oldData())) != length(unique(names(oldData())))) {
      errormsg_f("Error: There are more than one column with the same column name, 
               make sure the name of each column is different.")
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
    checkvaruniqold <- oldData() %>%
      select(input$key)
    checkvaruniqnew <- newData() %>%
      select(input$key)

    output$resultTable <- NULL
    output$downloadbtn <- NULL
    
    if (is.null(input$key)) {
      output$errorMsg <- renderText(
        HTML("<span style='color: red; font-size: 15px;'>No column has been selected.</span>")
      )
    }
    else if (nrow(oldData()) != nrow(unique(checkvaruniqold)) & nrow(newData()) != nrow(unique(checkvaruniqnew))) {
      output$errorMsg <- renderText(
        HTML("<span style='color: red; font-size: 15px;'>In both ORIGINAL and NEW file, there are more than one
             record with the same concatenated values from the selected columns </span>")
      )
    }
    else if (nrow(oldData()) != nrow(unique(checkvaruniqold))) {
      output$errorMsg <- renderText(
        HTML("<span style='color: red; font-size: 15px;'>In the ORIGINAL file, there are more than one
             record with the same concatenated values from the selected columns </span>")
      )
    }
    else if (nrow(newData()) != nrow(unique(checkvaruniqnew))) {
      output$errorMsg <- renderText(
        HTML("<span style='color: red; font-size: 15px;'>In the NEW file, there are more than one
             record with the same concatenated values from the selected columns </span>")
      )
    }
    else{
      #11Dec2025, added a snippet to handle scenario when all variables are selected as key variable to merge)
      olddf <- oldData()
      newdf <- newData()
      #replaced oldData()/newData() with olddf/newdf in previous snippet
      if (length(names(oldData()) == length(input$key))){
        olddf$tempvariable = "temp"
        newdf$tempvariable = "temp"
      }
      #end of snippet added on 11Dec2025
      output$errorMsg <- NULL
      merged <- merge(olddf, newdf, by = input$key, all = TRUE, suffixes = c(".old", ".new"))
      
      # Find common columns (excluding the key column)
      common_cols <- setdiff(names(olddf), input$key)
      
      # Create "Differences" column
      differences <- apply(merged, 1, function(row) {
        row_old <- row[paste0(common_cols, ".old")]
        row_new <- row[paste0(common_cols, ".new")]
        
        names(row_old) <- common_cols
        names(row_new) <- common_cols
        
        compare_row(row_old, row_new, common_cols)
      })
      
      merged$Differences <- differences
      output$resultTable <- renderTable(merged)
      output$downloadbtn <- renderUI(downloadButton("downloadData", "Download the output"))
      #11Dec2025, updated to remove temp variable in the output if all variables are selected to compare
      if ("tempvariable" %in% names(olddf)) {
        merged <- merged %>%
          select(-tempvariable.old, -tempvariable.new)
      }
      #end of 11Dec25 update
      output$downloadData <- downloadHandler(
        filename = function() {
          paste("merged_comparison_", Sys.Date(), ".csv", sep = "")
        },
        content = function(file) {
          write.csv(merged, file, row.names = FALSE)
        }
      )
    }
  })
}

shinyApp(ui = ui, server = server)
