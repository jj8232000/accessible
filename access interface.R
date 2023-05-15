# imports table from access into R

library(shiny)
library(RODBC)

ui <- fluidPage(
  titlePanel("MS Access Database Interface"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file", "Choose MS Access Database File (.mdb or .accdb)", accept = c(".mdb", ".accdb")),
      uiOutput("tableSelect"),
      actionButton("loadTableBtn", "Load Table"),
      tags$hr(),
      verbatimTextOutput("message")
    ),
    mainPanel(
      tableOutput("tableData")
    )
  )
)

server <- function(input, output, session) {
  
  # initialize reactive value storing table name
  table_var <- reactiveVal(NULL)
  
  # loads access db
  observeEvent(input$file, {
    file <- input$file$datapath
    conn <- odbcConnectAccess2007(file)
    tables <- sqlTables(conn)
    tables <- tables$TABLE_NAME
    tables <- tables[!grepl("^MSys", tables) & !grepl("^~", tables)]
    updateSelectInput(session, "table", choices = tables)
    odbcClose(conn)
  })
  
  # select table render
  output$tableSelect <- renderUI({
    selectInput("table", "Select Table", choices = NULL)
  })
  
  # update reactive value with table name
  observeEvent(input$table, {
    table_var(input$table)
  })
  
  # stores table data as a variable in the r environment
  observeEvent(input$loadTableBtn, {
    conn <- odbcConnectAccess2007(input$file$datapath)
    query <- paste("SELECT * FROM", table_var(), ";")
    data <- sqlQuery(conn, query)
    odbcClose(conn)
    assign(table_var(), data, envir = .GlobalEnv)
    output$tableData <- renderTable(data)
    output$message <- renderPrint({
      paste("The", table_var(), "table is loaded.")
    })
    showModal(modalDialog(
      title = "Table Loaded",
      paste(table_var(), "has been loaded into R."),
      easyClose = TRUE
    ))
  })
  
}

shinyApp(ui, server)
