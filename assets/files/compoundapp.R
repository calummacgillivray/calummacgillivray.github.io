library(shiny)
library(lubridate)
library(openxlsx)
library(DT)

ui <- fluidPage(
  titlePanel("Compound Interest by Period"),
  
  sidebarLayout(
    sidebarPanel(
      numericInput("initial_total", "Initial Total (£):", value = 1000, min = 0, step = 0.01),
      br(),
      actionButton("add_cp", "Add Compound Period", icon = icon("plus"), class = "btn-success"),
      actionButton("reset_app", "Reset App", icon = icon("redo"), class = "btn-danger"),
      br(), br(),
      actionButton("calc", "Calculate", class = "btn-primary"),
      br(), br(),
      downloadButton("export_excel", "Export Table to Excel")
    ),
    mainPanel(
      tags$div(id = "compound_periods_container"),
      br(),
      DTOutput("results_table"),
      br(),
      h4("Final End Total:"),
      textOutput("final_total")
    )
  )
)

server <- function(input, output, session) {
  cp_count <- reactiveVal(0)
  cp_interest_count <- reactiveValues()
  results_data <- reactiveVal(data.frame())
  
  # ---------- UI builders ----------
  compound_period_ui <- function(cp_num) {
    tags$div(
      id = paste0("compound_", cp_num),
      style = "border:1px solid #ccc; padding:15px; margin-bottom:15px; border-radius:8px; box-shadow:2px 2px 5px #eee;",
      h4(paste("Compound Period", cp_num)),
      fluidRow(
        column(6, dateInput(paste0("cp", cp_num, "_start"), "Start Date:", value = Sys.Date(), format = "dd/mm/yyyy")),
        column(6, dateInput(paste0("cp", cp_num, "_end"), "End Date:", value = Sys.Date() + 90, format = "dd/mm/yyyy"))
      ),
      numericInput(paste0("cp", cp_num, "_starting_interest"), "Starting Interest (%)", value = 0, step = 0.01),
      tags$div(id = paste0("cp", cp_num, "_interest_container")),
      actionButton(paste0("add_interest_", cp_num), "Add Interest Period", class = "btn-info btn-sm"),
      actionButton(paste0("remove_interest_", cp_num), "Remove Last Interest Period", class = "btn-warning btn-sm"),
      hr()
    )
  }
  
  add_compound_period <- function() {
    new_cp <- cp_count() + 1
    cp_count(new_cp)
    cp_interest_count[[paste0("cp", new_cp)]] <- 0
    
    insertUI(
      selector = "#compound_periods_container",
      where = "beforeEnd",
      ui = compound_period_ui(new_cp)
    )
    
    # ---------- Add/Remove interest observers ----------
    observeEvent(input[[paste0("add_interest_", new_cp)]], {
      req(!is.null(cp_interest_count[[paste0("cp", new_cp)]]))
      count <- cp_interest_count[[paste0("cp", new_cp)]] + 1
      cp_interest_count[[paste0("cp", new_cp)]] <- count
      
      insertUI(
        selector = paste0("#cp", new_cp, "_interest_container"),
        where = "beforeEnd",
        ui = fluidRow(
          style = "margin-bottom:10px;",
          column(6, dateInput(
            paste0("cp", new_cp, "_interest_", count, "_start"),
            paste("Interest Period", count, "Start:"),
            value = Sys.Date(), format = "dd/mm/yyyy"
          )),
          column(6, numericInput(
            paste0("cp", new_cp, "_interest_", count, "_rate"),
            "Interest (%)", value = 5, step = 0.01
          ))
        )
      )
    })
    
    observeEvent(input[[paste0("remove_interest_", new_cp)]], {
      req(!is.null(cp_interest_count[[paste0("cp", new_cp)]]))
      count <- cp_interest_count[[paste0("cp", new_cp)]]
      if (count > 0) {
        removeUI(selector = paste0("#cp", new_cp, "_interest_container > div:nth-child(", count, ")"))
        cp_interest_count[[paste0("cp", new_cp)]] <- count - 1
      }
    })
  }
  
  observe({
    if (cp_count() == 0) add_compound_period()
  })
  
  observeEvent(input$add_cp, { add_compound_period() })
  
  observeEvent(input$reset_app, {
    session$reload()
  })
  
  # ---------- Calculation ----------
  observeEvent(input$calc, {
    n_cp <- cp_count()
    overall_results <- data.frame(
      Compound = character(0),
      Interest.Period = character(0),
      Interest.Rate = numeric(0),
      Days = numeric(0),
      Period.Start = as.Date(character(0)),
      Period.End = as.Date(character(0)),
      Interest.Amount = numeric(0),
      End.Total = numeric(0),
      stringsAsFactors = FALSE
    )
    
    previous_cp_end_total <- input$initial_total
    
    for (cp in 1:n_cp) {
      cp_start <- input[[paste0("cp", cp, "_start")]]
      cp_end   <- input[[paste0("cp", cp, "_end")]]
      req(!is.null(cp_start), !is.null(cp_end))
      validate(need(cp_end >= cp_start, paste("Compound Period", cp, "end date must be on/after start date.")))
      
      starting_interest <- input[[paste0("cp", cp, "_starting_interest")]]
      n_ip <- cp_interest_count[[paste0("cp", cp)]]
      ip_list <- list(list(date = cp_start, interest = starting_interest))
      
      if (!is.null(n_ip) && n_ip > 0) {
        for (ip in 1:n_ip) {
          ip_start <- input[[paste0("cp", cp, "_interest_", ip, "_start")]]
          ip_rate  <- input[[paste0("cp", cp, "_interest_", ip, "_rate")]]
          if (!is.null(ip_start) && !is.null(ip_rate)) {
            ip_list <- append(ip_list, list(list(date = ip_start, interest = ip_rate)))
          }
        }
      }
      
      order_idx <- order(as.Date(sapply(ip_list, function(x) x$date)))
      ip_list <- ip_list[order_idx]
      dates <- as.Date(sapply(ip_list, function(x) x$date))
      rates <- as.numeric(sapply(ip_list, function(x) x$interest))
      dates <- c(dates, as.Date(cp_end) + 1)
      
      # ---------- Non-compounding interest per interest period ----------
      start_total <- previous_cp_end_total  # Start of compound period
      for (i in seq_along(rates)) {
        period_start <- dates[i]
        period_end   <- dates[i + 1] - 1
        days_in_period <- as.numeric(period_end - period_start + 1)
        days_in_year <- ifelse(leap_year(year(period_start)), 366, 365)
        
        # Interest calculated on the start of the compound period
        interest_amount <- round(previous_cp_end_total * (rates[i]/100 * (days_in_period/days_in_year)), 2)
        end_total_ip <- round(start_total + interest_amount, 2)
        
        overall_results <- rbind(
          overall_results,
          data.frame(
            Compound = paste("Compound Period", cp),
            Interest.Period = paste("Interest Period", i),
            Interest.Rate = paste0(rates[i], "%"),
            Days = days_in_period,
            Period.Start = format(period_start, "%d/%m/%Y"),
            Period.End = format(period_end, "%d/%m/%Y"),
            Interest.Amount = paste0("£", formatC(interest_amount, format="f", digits=2)),
            End.Total = paste0("£", formatC(end_total_ip, format="f", digits=2)),
            stringsAsFactors = FALSE
          )
        )
        start_total <- end_total_ip
      }
      
      previous_cp_end_total <- start_total
    }
    
    results_data(overall_results)
    
    # ---------- Render DT ----------
    output$results_table <- renderDT({
      df <- results_data()
      display_df <- data.frame()
      for (cp in unique(df$Compound)) {
        cp_df <- df[df$Compound == cp, ]
        display_df <- rbind(display_df, data.frame(Compound = cp, Interest.Period="", Interest.Rate="", Days="", Period.Start="", Period.End="", Interest.Amount="", End.Total="", stringsAsFactors = FALSE))
        display_df <- rbind(display_df, cp_df)
      }
      DT::datatable(display_df, rownames = FALSE, options = list(dom = 't')) %>%
        DT::formatStyle(
          'Compound',
          target = 'row',
          fontWeight = styleEqual(unique(display_df$Compound), rep('bold', length(unique(display_df$Compound)))),
          backgroundColor = styleEqual(unique(display_df$Compound), rep(c("#FDFDFD", "#FFFFFF"), length.out = length(unique(display_df$Compound))))
        )
    })
    
    output$final_total <- renderText({ tail(overall_results$End.Total, 1) })
  })
  
  # ---------- Export Excel ----------
  output$export_excel <- downloadHandler(
    filename = function() { paste0("compound_periods_", Sys.Date(), ".xlsx") },
    content = function(file) {
      df <- results_data()
      wb <- createWorkbook()
      addWorksheet(wb, "Compound Periods")
      
      display_df <- data.frame()
      for (cp in unique(df$Compound)) {
        cp_df <- df[df$Compound == cp, ]
        display_df <- rbind(display_df, data.frame(Compound = cp, Interest.Period="", Interest.Rate="", Days="", Period.Start="", Period.End="", Interest.Amount="", End.Total="", stringsAsFactors = FALSE))
        display_df <- rbind(display_df, cp_df)
      }
      
      writeData(wb, sheet = 1, display_df, startRow = 1, headerStyle = createStyle(textDecoration = "bold", fontSize = 12))
      
      for (i in 1:nrow(display_df)) {
        if (display_df$Interest.Period[i] == "") {
          addStyle(wb, sheet = 1, rows = i + 1, cols = 1:ncol(display_df), style = createStyle(fgFill = "#FDFDFD", textDecoration = "bold"), gridExpand = TRUE)
        }
      }
      
      currency_style <- createStyle(numFmt = "£#,##0.00")
      addStyle(wb, sheet = 1, rows = 2:(nrow(display_df)+1), cols = which(names(display_df) %in% c("Interest.Amount", "End.Total")), style = currency_style, gridExpand = TRUE)
      setColWidths(wb, sheet = 1, cols = 1:ncol(display_df), widths = "auto")
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
}

shinyApp(ui, server)
