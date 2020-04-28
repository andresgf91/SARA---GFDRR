require(openxlsx)

# all_grants <- raw_grants %>%
#   dplyr::rename("Available Balance (Uncommitted)" = `Available Balance USD`) %>% 
#   dplyr::mutate("Available Balance (Uncommitted)" =
#                   `Grant Amount USD` - (`Disbursements USD` + `Commitments USD`)) %>% 
#   dplyr::mutate("Real Transfers in" = `Transfer-in USD` - `Transfers-out in USD`) %>% 
#   dplyr::mutate("Not Yet Transferred" = `Grant Amount USD` - `Real Transfers in`) %>% 
#   select(-c("Transfer-in USD","Transfers-out in USD"))
# 
# all_grants <- all_grants %>%
#   mutate(`Fund Country Region Name`= Region) %>% 
#   select(-Region) %>% 
#   rename("Region Name"=`Fund Country Region Name`,
#          "Cumulative Disbursements USD" = `Disbursements USD`)
# 

active_grants <- report_grants %>%
  select(
    `Closing FY`,
    `Window #`,
    `Window Name`,
    `Trustee #`,
    `Child Fund #`,
    `Project ID`,
    `Child Fund Name`,
    `Execution Type`,
    `Child Fund Status`,
    `Child Fund TTL Name`,
    #`TTL Unit`,
    `Managing Unit`,
    `Lead GP/Global Theme`,
    VPU,
    `Region Name`,
    Country,
    `Activation Date`,
    `Closing Date`,
    `Grant Amount`,
    `Cumulative Disbursements`,
    `2020 Disbursements`,
    `PO Commitments`,
    `Real Transfers in`,
    `Not Yet Transferred`,
    `Remaining Available Balance`,
    `Months to Closing Date`,
    `Required Monthly Disbursement Rate`,
    `Disbursement Risk Level`
  )

active_grants <- active_grants %>%
  rename(`Available Balance (Uncommitted)`= `Remaining Available Balance`)

PMA_grants_excel <- active_grants %>% filter(`Child Fund #` %in% PMA_grants$Fund)

active_grants <- active_grants %>% 
  filter(!(`Child Fund #` %in% report_grants$`Child Fund #`[report_grants$PMA=='yes']))


active_grants$`Disbursement Risk Level` <- factor(active_grants$`Disbursement Risk Level` ,
                                           levels = c( "Closed (Grace Period)",
                                                       "Very High Risk",
                                                       "High Risk",
                                                       "Medium Risk",
                                                       "Low Risk"))

active_grants <- active_grants %>% arrange(`Disbursement Risk Level`)

active_grants$`Required Monthly Disbursement Rate`[active_grants$`Disbursement Risk Level`==
                                                     "Closed (Grace Period)"] <- NA


D.trustee <- trustee %>%
  select(`Donor Name`,Fund,`Fund Name`,`Fund TTL Name`,`TF End Disb Date`) 


names(D.trustee) <- c("Donor",
                    "Trustee",
                    "Trustee Fund Name",
                    "Trustee Fund TTL Name",
                    "End Disbursement Date")

D.trustee <- D.trustee %>% arrange(Donor)


#create excel workbook
wb <- loadWorkbook(grants_file)

#addtables 
names(wb) <-  "All Grants (SAP Data)"

addWorksheet(wb, "Active Grants")
writeDataTable(wb,"Active Grants", active_grants,startRow = 8,
               withFilter = TRUE,tableStyle = "TableStyleLight9")

addWorksheet(wb, "PMA Grants")
writeDataTable(wb,"PMA Grants", PMA_grants_excel,
               withFilter = TRUE,tableStyle = "TableStyleLight12")

addWorksheet(wb, "GFDRR TFs")
writeDataTable(wb,"GFDRR TFs", D.trustee,
               withFilter = TRUE,tableStyle = "TableStyleLight10")



#FORMATTING 
low_risk <- createStyle(fgFill ="#86F9B7",halign = "left")
medium_risk <- createStyle(fgFill ="#FFE285",halign = "left")
high_risk <- createStyle(fgFill ="#FD8D75",halign = "left")
very_high_risk <- createStyle(fgFill ="#F17979",halign = "left")
grace_period_style <- createStyle(fgFill ="#e0dede",halign = "left")

low_risk_rows <- which(active_grants$`Disbursement Risk Level`=="Low Risk")
medium_risk_rows <- which(active_grants$`Disbursement Risk Level`=="Medium Risk")
high_risk_rows <- which(active_grants$`Disbursement Risk Level`=="High Risk")
very_high_risk_rows <- which(active_grants$`Disbursement Risk Level`=="Very High Risk")
grace_period_rows <- which(active_grants$`Disbursement Risk Level`=="Closed (Grace Period)")


RISK.df.row <- 8

for (i in low_risk_rows){
  addStyle(wb,2,rows=i+RISK.df.row,cols=1:length(active_grants),style = low_risk)
}

for (i in medium_risk_rows){
  addStyle(wb,2,rows=i+RISK.df.row,cols=1:length(active_grants),style = medium_risk)
}

for (i in high_risk_rows){
  addStyle(wb,2,rows=i+RISK.df.row,cols=1:length(active_grants),style = high_risk)
}

for (i in very_high_risk_rows){
  addStyle(wb,2,rows=i+RISK.df.row,cols=1:length(active_grants),style = very_high_risk)
}

for (i in grace_period_rows){
  addStyle(wb,2,rows=i+RISK.df.row,cols=1:length(active_grants),style = grace_period_style)
}


legend_text <- c("Closed. These grants are under the grace period.",
                 "Very High Risk (>10% / month disbursement required)",
                 "High Risk ( > 5%-10% / month disbursement required)",
                 "Medium Risk (3 - 5% / month disbursement required)",
                 "Low Risk (<  3% / month disbursement required)")

sheet_header <- paste("GFDRR Active Grant List as of",report_data_date)


writeData(wb,2,sheet_header,startCol = 1,startRow = 1)
writeData(wb,2,"Disbursement Risk Legend",startCol = 6,startRow = 1)
writeData(wb,2,legend_text,startCol = 7,startRow = 2)

dollar_format <- createStyle(numFmt = "ACCOUNTING")
percent_format <- createStyle(numFmt = "0%")
date_format <- createStyle(numFmt = "mm/dd/yyyy")
num_format <- createStyle(numFmt = "NUMBER")
wrapped_text <- createStyle(wrapText = TRUE,valign = "top")
black_and_bold <- createStyle(fontColour = "#000000",textDecoration = 'bold')



addStyle(wb,2,rows=2,cols = 6, style = grace_period_style)
addStyle(wb,2,rows=3,cols = 6, style = very_high_risk)
addStyle(wb,2,rows=4,cols = 6, style = high_risk)
addStyle(wb,2,rows=5,cols = 6, style = medium_risk)
addStyle(wb,2,rows=6,cols = 6, style = low_risk)

# if(JAIME_preprocess == FALSE){
# 
# all_grants_row_range <- 1:nrow(all_grants)+1
# addStyle(wb,1,rows=all_grants_row_range,cols=25,style=date_format)   #activation date
# addStyle(wb,1,rows=all_grants_row_range,cols=26,style=date_format)   #closing date
# addStyle(wb,1,rows=all_grants_row_range,cols=27,style=dollar_format) #grant amount
# addStyle(wb,1,rows=all_grants_row_range,cols=28,style=dollar_format) #cumulative disbursements
# addStyle(wb,1,rows=all_grants_row_range,cols=29,style=dollar_format) #2020 disbursements
# addStyle(wb,1,rows=all_grants_row_range,cols=30,style=dollar_format) #commitments
# addStyle(wb,1,rows=all_grants_row_range,cols=31,style=dollar_format) #available balance
# addStyle(wb,1,rows=all_grants_row_range,cols=32,style=dollar_format) #real- transfer-in
# addStyle(wb,1,rows=all_grants_row_range,cols=33,style=dollar_format) #Not yet Transferred
# }

active_grants_row_range <- 9:(nrow(active_grants)+8)
addStyle(wb,2,rows=active_grants_row_range,cols=16,style=date_format,stack = TRUE)   #activation date
addStyle(wb,2,rows=active_grants_row_range,cols=17,style=date_format,stack = TRUE)   #closing date
addStyle(wb,2,rows=active_grants_row_range,cols=18,style=dollar_format,stack = TRUE) #grant amount
addStyle(wb,2,rows=active_grants_row_range,cols=19,style=dollar_format,stack = TRUE) #cumulative disbursements
addStyle(wb,2,rows=active_grants_row_range,cols=20,style=dollar_format,stack = TRUE) #2020 disbursements
addStyle(wb,2,rows=active_grants_row_range,cols=21,style=dollar_format,stack = TRUE) #commitments
addStyle(wb,2,rows=active_grants_row_range,cols=22,style=dollar_format,stack = TRUE) #real- transfer-in
addStyle(wb,2,rows=active_grants_row_range,cols=23,style=dollar_format,stack = TRUE) #Not yet Transferred
addStyle(wb,2,rows=active_grants_row_range,cols=24,style=dollar_format,stack = TRUE) #available balance
addStyle(wb,2,rows=active_grants_row_range,cols=26,style=percent_format,stack = TRUE) #required monthly disbursement rate

addStyle(wb,2,rows = 8,cols = 1:ncol(active_grants),style = wrapped_text,stack = TRUE)
setColWidths(wb,2,cols = 7,widths = 30)
setColWidths(wb,2, cols = 10,widths = 12)
setColWidths(wb,2,cols = 16:24,widths = 15)
setColWidths(wb,2,cols = 25,widths = 12)
setColWidths(wb,2,cols = 26,widths = 12)
setColWidths(wb,2,cols = 27,widths = 20)



PMA_grants_row_range <- 1:nrow(PMA_grants_excel)+1
addStyle(wb,3,rows=PMA_grants_row_range,cols=16,style=date_format,stack = TRUE)   #activation date
addStyle(wb,3,rows=PMA_grants_row_range,cols=17,style=date_format,stack = TRUE)   #closing date
addStyle(wb,3,rows=PMA_grants_row_range,cols=18,style=dollar_format,stack = TRUE) #grant amount
addStyle(wb,3,rows=PMA_grants_row_range,cols=19,style=dollar_format,stack = TRUE) #cumulative disbursements
addStyle(wb,3,rows=PMA_grants_row_range,cols=20,style=dollar_format,stack = TRUE) #2020 disbursements
addStyle(wb,3,rows=PMA_grants_row_range,cols=21,style=dollar_format,stack = TRUE) #commitments
addStyle(wb,3,rows=PMA_grants_row_range,cols=22,style=dollar_format,stack = TRUE) #real- transfer-in
addStyle(wb,3,rows=PMA_grants_row_range,cols=23,style=dollar_format,stack = TRUE) #Not yet Transferred
addStyle(wb,3,rows=PMA_grants_row_range,cols=24,style=dollar_format,stack = TRUE) #available balance
addStyle(wb,3,rows=PMA_grants_row_range,cols=26,style=percent_format,stack = TRUE)
addStyle(wb,3,rows = 1,cols = 1:ncol(PMA_grants_excel),style = wrapped_text,stack = TRUE)


addStyle(wb,4,rows=1:nrow(D.trustee)+1,cols=5,style=date_format,stack = TRUE)
setColWidths(wb,4, cols = 1:ncol(D.trustee), widths = "auto")


