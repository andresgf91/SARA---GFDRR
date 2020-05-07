#
# This is the user-interface definition of the Shiny web application. You can
# run the application by clicking 'Run App' above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/ 
#

# LOAD PACKAGES -----------------------------------------------------------

library(shinydashboard)
library(shiny)
library(shinyWidgets)
library(shinydashboardPlus)
library(readxl)
library(pander)
library(scales)
library(lubridate)
library(ggplot2)
library(hrbrthemes)
library(dplyr)
library(tidyr)
library(viridis)
library(shinyBS)
library(plotly)
library(streamgraph)
library(DT)
library(shinythemes)
library(shinyjs)
library(waiter)
library(bsplus)
library(tippy)

# BUILD USER INTERFACE ----------------------------------------------------

## Header ---------------
Header <- dashboardHeaderPlus(
  title =
    tagList(
      span(class = "logo-lg", 'Menu'),
      span(class = "logo-mini", "Menu")
    )
  ,
  dropdownMenuOutput("date_data_updated_message"),
  # disable = F,
  enable_rightsidebar = TRUE,
  rightSidebarIcon = "filter",
  left_menu =  tagList(
    img(
      class = "logo-lg",
      src = 'SARA Logo_canva_black.png',
      height = '45',
      width = '190'
    ),
    img(
      class = "logo-lg",
      src = 'GFDRR_BW_logo.png',
      height = '45',
      width = '190'
    )
  )
)
## Sidebar ---------------

    
secretariat_view <-   menuItem(
  "Portfolio Overview",
  tabName = "overview",
  icon = icon("dashboard"),
  selected = T
)

info <- menuItem("Additional Information",
                 tabName = "admin_info",
                 icon = icon("info"))

parent_trust_fund_view <-   menuItem(
  "Parent Trust Fund",
  tabName = "parent_tf",
  icon = icon("university"),
  selected = F
)

regions_view <-  menuItem(
  "Grant Portfolio",
  tabName = "regions",
  icon = icon("briefcase"),
  selected = F
)

TTL_grant_detail <-  menuItem(
  "TTL/Grant Detail",
  menuSubItem(
    "Grant View",
    tabName = "grant_dash",
    icon = icon("dashboard")
  ),
  menuSubItem("TTL View",
              tabName = "TTL_dashboard",
              icon = icon("stream"))
)

reports_tab <- menuItem(
  "Download Reports",
  tabName = "reports",
  icon = icon("file-download"),
  selected = F
)


Sidebar <- dashboardSidebar(
  #collapsed = TRUE,
  sidebarMenu(
    secretariat_view,
    parent_trust_fund_view,
    regions_view,
    info,
    reports_tab,
    # TTL_grant_detail,
    id = 'nav'
  )
)

## tab.1 (Overview)---------------
tab.1 <-  tabItem(tabName = "overview",
                    theme = shinytheme("readable"),
                     titlePanel('Portfolio Overview'),
                  panel(textOutput("caption.1")),
                    fluidRow(
                      column(
                        width = 2,
                        #tags$div(title="This is an example of how we could add text to the different elements in the dashboard",
                                 valueBoxOutput("total_contributions", width = NULL),
                        #),
                      valueBoxOutput("total_received", width = NULL), #%>% 
                        #bs_embed_tooltip(title = "Example of where we could add text to explain the Indicators", placement = 'right'), 
                        valueBoxOutput("total_unpaid", width = NULL), #%>% 
                       # bs_embed_tooltip(title = "Testtt222"),
                        valueBoxOutput("total_active_portfolio", width = NULL),
                        valueBoxOutput("total_uncommitted_balance", width = NULL),
                        valueBoxOutput("closing<12", width = NULL)),
                      column(
                        width = 4,
                       
                        boxPlus(id="myid",
                          plotlyOutput("funding_region", height = "260px"),
                                title='Funding by Region',
                                background = "blue",
                                #enable_label = T,
                                #label_text = tippy("i","This is an example of how we could add more information to the graphs so the user #understand what it represents",width = '90',height = '120px'),
                                width = NULL,
                                collapsible = TRUE,
                                closable = F,
                                collapsed = F,
                          enable_dropdown = F),
                       # tippy_this("myid", "Tooltip",placement='right'),
                        boxPlus(
                          plotlyOutput("funding_GP", height = "260px"),
                                title='Funding by Global Practice',
                                background = "blue",
                                enable_label = T,
                                label_text = NULL,
                                width = NULL,
                                collapsible = TRUE,
                                closable = FALSE,
                                collapsed = F),
                        boxPlus(
                          plotlyOutput("n_grants_region", height = 260),
                          title='Grants by Region',
                          background = "blue",
                          enable_label = T,
                          label_text = NULL,
                          width = NULL,
                          collapsible = TRUE,
                          closable = F,
                          collapsed = F)
                       # plotlyOutput("funding_GP", height = 550)
                      ),
                      column(
                        width = 6,
                        fluidRow(
                          column(width = 6,
                                 boxPlus(title = "Active grants by activation FY",
                                plotlyOutput("pie_grants_by_FY",
                                             height = "260px"),
                                background = "blue",
                                enable_label = T,
                                label_text = NULL,
                                width = NULL,
                                collapsible = TRUE,
                                closable = F)),
                          column(width = 6,
                                 boxPlus(
                                 plotlyOutput("elpie",
                                              height = "260px"),
                                 #),
                          title = 'Active Portfolio',
                          background = "blue",
                          enable_label = T,
                          label_text = NULL,
                          width = NULL,
                          collapsible = TRUE,
                          closable = F
                        ))),
                        
                        boxPlus(
                          plotlyOutput("plot1",
                                       height = "600px"),
                                title = "Contributions by Parent Fund",
                                width = NULL,
                          background = "blue",
                          enable_label = T,
                          label_text = NULL,
                          collapsible = TRUE,
                          closable = F)
                        #plotlyOutput("overview_progress_GG", height = 90)),
                      )
                    )
                          
                    
                    
    )



# tab 1.3 (additional info)------
tab.1.3 <-  tabItem(tabName = "admin_info",
                      titlePanel('Additional Information'),
                    panel(textOutput("caption.4")),
                    
                      tabsetPanel(
                        type = 'pills',
                        tabPanel("Glossary",
                                 tableOutput("glossary_1")),
                        tabPanel("Parent Fund Info",
                                 tableOutput("trustee_name_TTL")),
                        tabPanel(
                          "Donor Contributions",
                          tableOutput("donor_contributions"),
                          plotlyOutput("donor_contributions_GG")
                        )#,
                        #tabPanel(
                         # title = "RETF grants overview",
                          #fluidRow(
                           # valueBoxOutput("RETF_n_grants_A"),
                            #valueBoxOutput("RETF_$_grants_A")
                          #),
                          #fluidPage(tabsetPanel(
                           # tabPanel(title = "Trustees",
                            #         plotlyOutput("RETF_trustees_A_pie", height = 450)),
                            #tabPanel(title = "Regions",
                             #        plotlyOutput("RETF_region_A_pie",height = 450))
                         # ))
                       # )
                      )
                    )

## tab.2 (Parent Trustfund View) ----------------------
tab.2 <- tabItem(tabName = "parent_tf",
                 fluidPage(
                   theme = "light",
                   fluidRow(
                     column(width=10,
                            titlePanel("Parent Trust Fund View")),
                     column(width=2,
                            actionBttn('opensidebar.1',
                                       "Click to filter current view",
                                       size='sm',
                                       icon=icon("filter")))),
                   panel(textOutput("caption.2")),
                   tabsetPanel(
                   tabPanel("Main View",  
                   
                   fluidRow(column(
                     width = 3,
                     boxPlus(
                       title = "",
                       closable = F,
                       enable_label = T,
                       label_text = 'Parent Fund(s)',
                       textOutput('trustee_name'),
                       width = NULL,
                       background = "navy"
                     )
                   ),
                   column(
                     width = 9,
                     boxPlus(
                       title = "",
                       closable = F,
                       enable_label = T,
                       label_text = 'Donor Agencies',
                       textOutput('trustee_contribution_agency'),
                       width = 6,
                       background = 'navy'
                     ),
                     boxPlus(
                       title = "",
                       closable = F,
                       enable_label = T,
                       label_text = 'Fund Manager(s)',
                       textOutput('TTL_name'),
                       width = 6,
                       background = "navy"
                     )
                   )),
                   fluidRow(
                     column(
                       width = 3,
                       valueBoxOutput("fund_balance", width = NULL),
                       valueBoxOutput("trustee_closing_in_months", width = NULL),
                       valueBoxOutput("trustee_active_grants", width = NULL),
                       valueBoxOutput("trustee_grants_closing_3", width = NULL)
                     ),
                     column(
                       width = 9,
                       boxPlus(
                         plotlyOutput("trustee_received_unpaid_pie", height = 260),
                         title = 'Parent Fund(s) Balance',
                         background = "blue",
                         enable_label = T,
                         label_text = NULL,
                         width = 6,
                         collapsible = TRUE,
                         closable = F
                       ),
                       boxPlus(
                         plotlyOutput("trustee_dis_GG", height = 260),
                         title = 'Active Portfolio Disbursement Summary',
                         background = "blue",
                         enable_label = T,
                         label_text = NULL,
                         width = 6,
                         collapsible = TRUE,
                         closable = F
                       ),
                       boxPlus(
                         plotlyOutput("trustee_region_n_grants_GG", height = 260),
                         title = 'Grants by Region',
                         background = "blue",
                         enable_label = T,
                         label_text = NULL,
                         width = 6,
                         collapsible = TRUE,
                         closable = F
                       ),
                       boxPlus(
                         plotlyOutput("trustee_region_GG", height = 260),
                         title = 'Funding per Region',
                         background = "blue",
                         enable_label = T,
                         label_text = NULL,
                         width = 6,
                         collapsible = TRUE,
                         closable = F
                       )
                     )
                   )),
                   tabPanel("Table View",
                   fluidRow(
                     boxPlus(
                       dataTableOutput("trustee_countries_DT"),
                       title = 'Country Summary Table',
                       background = NULL,
                       enable_label = T,
                       label_text = NULL,
                       width = NULL,
                       collapsible = TRUE,
                       closable = F,
                       collapsed = F,
                       footer = downloadBttn('Dload_country_summary_table')
                     )
                     
                   ))
                   )
                 ))


                      #valueBoxOutput("trustee_received",width = NULL),
                             # valueBoxOutput("trustee_unpaid",width = NULL)),
  


## tab.3 (Grant Portfolio View) ---------------

tab.3 <-  tabItem(
  tabName = "regions",
  class = 'active',
  fluidRow(
    column(width=10,
           titlePanel("Grant Portfolio")),
    column(width=2,
           actionBttn('opensidebar.2',
                      "Click to filter current view",
                      size='sm',
                      icon=icon("filter")))),
  panel(textOutput("caption.3")),
  tabsetPanel(
    tabPanel("Main View",
  fluidRow(
    column(
      width = 3,
      valueBoxOutput(outputId = "focal_active_grants", width = 12),
      valueBoxOutput(outputId = "focal_active_funds", width = 12),
      valueBoxOutput("low_risk", width = 6),
      valueBoxOutput("medium_risk", width = 6),
      valueBoxOutput("high_risk", width = 6),
      valueBoxOutput("very_high_risk", width = 6),
      valueBoxOutput("grace_period_grants", width = 6),
      valueBoxOutput("grants_active_this_FY", width = 6),
      valueBoxOutput(outputId = "focal_grants_closing_3", width = 6),
      valueBoxOutput(outputId = "focal_grants_active_3_zero_dis", width = 6),
      valueBoxOutput(outputId = "region_grants_may_need_transfer", width = 6),
      valueBoxOutput(outputId = "region_grants_active_no_transfer", width = 6)
    ),
    
    column(width = 9, fluidRow(
      boxPlus(
        plotlyOutput("elpie2",
                     height = "260px"),
        title = 'Portfolio Funds Summary',
        background = "navy",
        enable_label = T,
        label_text = NULL,
        width = 6,
        collapsible = TRUE,
        closable = F,
        collapsed = F
      ),
      boxPlus(
        plotlyOutput(outputId = "region_GP_GG",
                     height = "260px"),
        title = 'Funding by GP/Global Theme',
        background = "navy",
        enable_label = T,
        label_text = NULL,
        width = 6,
        collapsible = TRUE,
        closable = F,
        collapsed = F
      ),
      boxPlus(
        plotlyOutput(outputId = "disbursement_risk_GG",
                     height = "260px"),
        title = 'Grants Disbursement Risk',
        background = "blue",
        enable_label = T,
        label_text = NULL,
        width = 12,
        collapsible = TRUE,
        closable = F,
        collapsed = F
      ),
      
      boxPlus(
        plotlyOutput(outputId = "focal_region_n_grants_GG",
                     height = "420px"),
        title = 'Active Portfolio by Parent Fund',
        background = "light-blue",
        enable_label = T,
        label_text = NULL,
        width = 12,
        collapsible = TRUE,
        closable = F,
        collapsed = F
      )
    )))),
  tabPanel("Table View",
    fluidRow(
      boxPlus(
        solidHeader = T,
        DT::dataTableOutput(outputId = "region_summary_grants_table"),
        title = 'Summary Table',
        background = NULL,
        enable_label = T,
        label_text = NULL,
        width = 12,
        collapsible = TRUE,
        closable = F,
        collapsed = F,
        footer = downloadBttn('Dload_region_summary_grants_table')
      ) ,
      boxPlus(
        solidHeader = T,
        DT::dataTableOutput(outputId = "region_countries_grants_table"),
        title = 'Countries Summary Table',
        background = NULL,
        enable_label = T,
        label_text = NULL,
        width = 12,
        collapsible = TRUE,
        closable = F,
        collapsed = F,
        footer = downloadBttn('Dload_region_countries_grants_table')
      ) ,
     boxPlus(
        solidHeader = T,
        DT::dataTableOutput(outputId = "region_funding_source_grants_table"),
        title = 'Funding source table',
        background = NULL,
        enable_label = T,
        label_text = NULL,
        width = 12,
        collapsible = TRUE,
        closable = F,
        collapsed = F,
        footer = downloadBttn('Dload_region_funding_source_grants_table')
      )
    )
    )
    
  )
)
                  

## tab.4 (PMA View) ---------------

PMA.tab <- tabItem(tabName = "PMA",
                   titlePanel("Program Management and Administration"),
                             column(width=4,fluidRow(
                     infoBoxOutput(outputId = 'resources_available',width = NULL),
                     infoBoxOutput(outputId = 'PMA_grants_n',width = NULL),
                     infoBoxOutput(outputId = "current_quarter",width = NULL),
                     infoBoxOutput(outputId = "next_quarter",width = NULL),
                            progressBar(id = 'quarter_spent',
                                 value = 0,
                                 display_pct = T,
                                 title = "% Spent",size = NULL))),
                     column(width=8,
                     tabsetPanel(tabPanel(title= "PMA Quartely",
                     plotlyOutput(outputId = "PMA_chart_1")),
                     tabPanel(title= "PMA Streamgraph",
                     streamgraphOutput(outputId = "streamgraph",
                                       height="350px",
                                       width="800px")))))



## tab. 5 GRANT DASHBOARD -----------

tab.5 <- tabItem(
  tabName = "grant_dash",
    titlePanel("Grant Summary"),
    fluidRow(
      column(width = 3,
      box(width = NULL,
        textInput("child_TF_num",
                "Grant Number",
                value = NULL,
                placeholder = "Enter grant number (e.g., TF018353)")),
      box(title = 'Grant Name',
          textOutput('grant_name'),
          width = NULL),
      box(title = 'TTL Name',
          textOutput('grant_TTL'),
          width = NULL),
      box(title = 'Country',
          textOutput('grant_country'),
          width = NULL),
      box(title = 'Region',
          textOutput('grant_region'),
          width = NULL),
      box(title = 'Unit',
          textOutput('grant_unit'),
          width = NULL)),
      column(width=3,
             valueBoxOutput("single_grant_amount",width = NULL),
             valueBoxOutput("single_grant_remaining_bal",width = NULL),
             valueBoxOutput("single_grant_m_active",width = NULL),
             valueBoxOutput("single_grant_m_disrate",width = NULL),
             valueBoxOutput("single_grant_m_to_close",width = NULL),
             valueBoxOutput("single_grant_m_req_disrate",width = NULL))
      #,
     # box(
         # plotlyOutput("grant_expense_GG", height = 400)),
     # box(width = 4,collapsible = T,
       # tableOutput("expense_table"))
    )
  )




## TAB 6. TTL Dashboard ----------
tab.6 <- tabItem(
  tabName = "TTL_dashboard",
    titlePanel("TTL Summary"),
    fluidRow(
      column(width = 3,
             box(width = NULL,
                 textInput("TTL_upi",
                           "TTL UPI",
                           value = NULL,
                           placeholder = "Enter UPI")),
             box(title = 'TTL Name',
                 textOutput('TTL_name_dash'),
                 width = NULL),
             box(title = 'TTL Unit',
                 textOutput('TTL_unit_dash'),
                 width = NULL)),
      column(width=3, valueBoxOutput("TTL_grants_active",width = NULL),
             valueBoxOutput("TTL_total_grant_amount",width = NULL),
             valueBoxOutput("TTL_total_remaining_bal",width = NULL)),
      box(
        plotlyOutput("TTL_balance_GG", height = 400)),
      box(width = 4,collapsible = T,
          tableOutput("TTL_grant_table"))
    )
  )


# Tab 7. Reports Download Tab --------

tab.reports <- tabItem(
  tabName = 'reports',
  h1("Download Reports/Data"),
  panel(textOutput("caption.5")),
  fluidRow(
    column(width=2,
           selectInput(inputId = "report_type",label=NULL,
                       choices= c("Summary Report",
                                  "Disbursement Risk Report",
                                  "Pivot Table Report",
                                  "Combined Report (Master)",
                                  "SAP Data (Source Data)"),
                       selectize=TRUE),
           panel(heading = "Report Description",textOutput("report_description"))),
    column(width=5,
           conditionalPanel(
             condition = "input.report_type == 'Summary Report'",
             boxPad(color = 'blue',
                    selectInput(inputId = "summary_fund_status",width = NULL,
                                label="Child Fund Status:",
                                choices= c("ACTV","PEND"),
                                multiple = T,
                                selectize = T,
                                selected =  c("ACTV","PEND")),
                    selectInput(inputId = "summary_region",
                                label="Region Name(s):",
                                choices= sort(unique(grants$Region)),
                                multiple = T,
                                selectize = T,
                                selected =  sort(unique(grants$Region))),
                    selectInput(inputId = "summary_trustee",
                                label="Trustee(s):",
                                choices= sort(unique(grants$temp.name)),
                                multiple = T,
                                selectize = T,
                                selected =  sort(unique(grants$temp.name)))
             )
           ),
           conditionalPanel(
             condition = "input.report_type == 'Disbursement Risk Report'",
             boxPad(color = 'blue',
                    selectInput(inputId = "risk_fund_status",width = NULL,
                                label="Child Fund Status:",
                                choices= c("ACTV","PEND"),
                                multiple = T,
                                selectize = T,
                                selected =  c("ACTV","PEND")),
                    selectInput(inputId = "risk_region",
                                label="Region Name(s):",
                                choices= c("ALL",sort(unique(grants$Region))),
                                multiple = F,
                                selectize = T,
                                selected =  "ALL"),
                    selectInput(inputId = "risk_trustee",
                                label="Trustee(s):",
                                choices= sort(unique(grants$temp.name)),
                                multiple = T,
                                selectize = T,
                                selected =  sort(unique(grants$temp.name)))
             )
           ),
           conditionalPanel(
             condition = "input.report_type == 'Pivot Table Report'",
             boxPad(color = 'blue',
                    textOutput("caption.edit_pivot"),
                    checkboxGroupButtons( 'pivot_fund_status',
                                          "Selected fund status:",
                                          choices = c("ACTV","PEND"),
                                          selected = c("ACTV","PEND"),
                                          size = "sm",
                                          justified = F,
                                          status = "primary",
                                          individual = T,
                                          width='100%',
                                          checkIcon = list(yes = icon("ok", lib = "glyphicon"))),
                    # selectInput(inputId = "pivot_fund_status",width = '100%',
                    #             label="Child Fund Status:",
                    #             choices= c("ACTV","PEND"),
                    #             multiple = T,
                    #             selectize = T,
                    #             selected =  c("ACTV","PEND")),
                    checkboxGroupButtons( 'pivot_region',
                                          "Selected Regions:",
                                          choices = sort(unique(grants$Region)),
                                          selected = sort(unique(grants$Region)),
                                          size = "sm",
                                          justified = F,
                                          status = "primary",
                                          individual = T,
                                          width='100%',
                                          checkIcon = list(yes = icon("ok", lib = "glyphicon"))),
                    # selectInput(inputId = "pivot_region",width = '100%',
                    #             label="Region Name(s):",
                    #             choices= sort(unique(grants$Region)),
                    #             multiple = T,
                    #             selectize = T,
                    #             selected =  sort(unique(grants$Region))),
                    # selectInput(inputId = "pivot_trustee",width = '100%',
                    #             label="Trustee(s):",
                    #             choices= sort(unique(grants$temp.name)),
                    #             multiple = T,
                    #             selectize = T,
                    #             selected =  sort(unique(grants$temp.name))),
                    checkboxGroupButtons( 'pivot_trustee',
                                          "Selected Trustees (Parent Funds):",
                                          choices = sort(unique(grants$temp.name)),
                                          selected = sort(unique(grants$temp.name)),
                                          size = "sm",
                                          justified = F,
                                          status = "primary",
                                          individual = T,
                                          width='100%',
                                          checkIcon = list(yes = icon("ok", lib = "glyphicon"))),
                    # selectInput(inputId = "pivot_exec_type",width = '100%',
                    #             label="Execution Type:",
                    #             choices= sort(unique(grants$`DF Execution Type`)),
                    #             multiple = T,
                    #             selectize = T,
                    #             selected =  sort(unique(grants$`DF Execution Type`))),
                    checkboxGroupButtons( 'pivot_exec_type',
                                          "Selected execution types:",
                                          choices = sort(unique(grants$`DF Execution Type`)),
                                          selected = sort(unique(grants$`DF Execution Type`)),
                                          size = "sm",
                                          justified = F,
                                          status = "primary",
                                          individual = T,
                                          width='100%',
                                          checkIcon = list(yes = icon("ok", lib = "glyphicon"))),
                    # selectInput(inputId = "pivot_GPURL_binary",width = '100%',
                    #             label="GPURL Binary",
                    #             choices= sort(unique(grants$GPURL_binary)),
                    #             multiple = T,
                    #             selectize = T,
                    #             selected =  sort(unique(grants$GPURL_binary))),
                    checkboxGroupButtons( 'pivot_GP',
                                          "Selected GP/Global Themes:",
                                          choices = sort(unique(grants$`Lead GP/Global Themes`)),
                                          selected = sort(unique(grants$`Lead GP/Global Themes`)),
                                          size = "sm",
                                          justified = F,
                                          status = "primary",
                                          individual = T,
                                          width='100%',
                                          checkIcon = list(yes = icon("ok", lib = "glyphicon")))
                    #selectInput(inputId = "pivot_GP",width = '100%',
                               # label="Lead GP/Global Themes",
                               # choices= sort(unique(grants$`Lead GP/Global Themes`)),
                               # multiple = T,
                               # selectize = T,
                               # selected =  sort(unique(grants$`Lead GP/Global Themes`)))
           )
    ),
    conditionalPanel(condition = "input.report_type == 'Summary Report'",
                     downloadButton("Download_summary_report.xlsx",
                                    label = 'Download Report')),
    conditionalPanel(condition = "input.report_type == 'Disbursement Risk Report'",
                     downloadButton("Download_risk_report.xlsx",
                                    label = 'Download Report')),
    conditionalPanel(condition = "input.report_type == 'Combined Report (Master)'",
                     downloadButton("Download_master_report.xlsx",
                                    label = 'Download Report')),
    conditionalPanel(condition = "input.report_type == 'Pivot Table Report'",
                     downloadButton("Download_pivot_report.xlsx",
                                    label = 'Download Report')),
    conditionalPanel(condition = "input.report_type == 'SAP Data (Source Data)'",
                     downloadButton("Download_source_data.xlsx",
                                    label = 'Download Data'))
  )
)
)


## BODY ---------------
Body <- dashboardBody(
tags$head(tags$style(HTML('
  .navbar-custom-menu>.navbar-nav>li>.dropdown-menu {
  width:600px;
  }
    /* body */
    .content-wrapper, .right-side {
    background-color: #FFFFFF;
    }
    
    .small-box {height: 150px}
    
    .wrapper{
    overflow-y: hidden;
}

#pledged{color: white;
font-size: 22px;
font-style: italic;


}

 #report_description {font-size:20px;
               display:block; }

  '))),
tabItems(tab.1,tab.3,tab.2,tab.1.3,tab.reports))
#tab.2,tab.3,PMA.tab,tab.1.3,tab.5,tab.6))




# UI ------
ui <- dashboardPagePlus(title = "BETA-SARA-GFDRR",
                        collapse_sidebar = T,
                        useShinyjs(),
                        header = Header,
                        sidebar = Sidebar,
                        body = Body,
                        rightsidebar =  rightSidebar(
                          background = "dark",
                          uiOutput("side_bar"),
                          title = "Right Sidebar"),
                        skin = "black")

