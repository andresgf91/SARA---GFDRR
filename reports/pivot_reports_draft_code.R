
library(pivottabler)
require(openxlsx)


data <- report_grants

data <- data %>% mutate(subset=ifelse(`Trustee #` %in% current_trustee_subset,"yes","no"))

data$subset_available_bal <- ifelse(data$subset=='yes',
                                    data$`Remaining Available Balance`,0)

data$subset_po_commitments <- ifelse(data$subset=='yes',
                                     data$`PO Commitments`,0)

#SUMMARY WITH ALL DATA
pt.8 <- PivotTable$new(argumentCheckMode='minimal')
pt.8$addData(data)
pt.8$addRowDataGroups("GPURL_binary") 
pt.8$defineCalculation(calculationName="Grant Count",
                       summariseExpression="n()")
pt.8$defineCalculation(calculationName="Total $",
                       summariseExpression="sum(`Grant Amount`)")
pt.8$defineCalculation(calculationName="Total Available Balance (Uncommitted)",
                       summariseExpression="sum(`Remaining Available Balance`)")
pt.8$defineCalculation(calculationName="Percen Available Balance (Uncommitted)",
                       summariseExpression="mean(percent_unaccounted)/100")
pt.8$defineCalculation(calculationName="Total Available Balance (Uncommitted) under ACP-EU NDRR; Core MDTF; Japan Program Phase I SDTF; Parallel MDTF",summariseExpression="sum(subset_available_bal)")
pt.8$evaluatePivot()
df.8 <- pt.8$asDataFrame() %>% t() %>% as.data.frame()



#SUBSET OF SUBSET OF USER SELECTED DATA 

data <- data %>%
  filter(
    `Child Fund Status` %in% input$pivot_fund_status,
    `Region Name` %in% input$pivot_region,
    `Trustee Fund Name` %in% input$pivot_trustee,
    `Execution Type` %in% input$pivot_exec_type,
    GPURL_binary %in% input$pivot_GPURL_binary,
    `Lead GP/Global Theme` %in% input$pivot_GP
  )


#Grant Details (by Region and Country)
pt.1 <- PivotTable$new(argumentCheckMode='minimal')
pt.1$addData(data)
pt.1$addRowDataGroups("Region Name",outlineBefore=list(isEmpty=FALSE), outlineTotal=list(isEmpty=FALSE))
pt.1$addRowDataGroups("Country") 
pt.1$defineCalculation(calculationName="#Grants", summariseExpression="n()")
pt.1$defineCalculation(calculationName="$Total Grant Amount", summariseExpression="sum(`Grant Amount`)")
pt.1$defineCalculation(calculationName="Total Available Balance (Uncommitted)",
                       summariseExpression="sum(`Remaining Available Balance`)")
pt.1$evaluatePivot()


pt.2 <- PivotTable$new(argumentCheckMode='minimal')
pt.2$addData(data)
pt.2$addRowDataGroups("Trustee Fund Name") 
pt.2$defineCalculation(calculationName="#Grants", summariseExpression="n()")
pt.2$defineCalculation(calculationName="$Total Grant Amount", summariseExpression="sum(`Grant Amount`)")
pt.2$defineCalculation(calculationName="Total Available Balance (Uncommitted)",
                       summariseExpression="sum(`Remaining Available Balance`)")
pt.2$evaluatePivot()

pt.3 <- PivotTable$new(argumentCheckMode='minimal')
pt.3$addData(data)
pt.3$addRowDataGroups("Lead GP/Global Theme") 
pt.3$defineCalculation(calculationName="#Grants", summariseExpression="n()")
pt.3$defineCalculation(calculationName="$Total Grant Amount", summariseExpression="sum(`Grant Amount`)")
pt.3$defineCalculation(calculationName="Total Available Balance (Uncommitted)",
                       summariseExpression="sum(`Remaining Available Balance`)")
pt.3$evaluatePivot()

#Grant Details (by Region and Country)
pt.4 <- PivotTable$new(argumentCheckMode='minimal')
pt.4$addData(data)
pt.4$addColumnDataGroups("Disbursement Risk Level")
pt.4$addRowDataGroups("Region Name",outlineBefore=list(isEmpty=FALSE),outlineTotal=list(isEmpty=FALSE))
pt.4$addRowDataGroups("Country") 
pt.4$defineCalculation(calculationName="Available Balance (Uncommitted)",
                       summariseExpression="sum(`Remaining Available Balance`)")
pt.4$evaluatePivot()


pt.5 <- PivotTable$new(argumentCheckMode='minimal')
pt.5$addData(data)
pt.5$addColumnDataGroups("Disbursement Risk Level")
pt.5$addRowDataGroups("Trustee Fund Name") 
pt.5$defineCalculation(calculationName="Available Balance (Uncommitted)",
                       summariseExpression="sum(`Remaining Available Balance`)")
pt.5$evaluatePivot()


pt.6 <- PivotTable$new(argumentCheckMode='minimal')
pt.6$addData(data)
pt.6$addColumnDataGroups("Disbursement Risk Level")
pt.6$addRowDataGroups("Lead GP/Global Theme") 
pt.6$defineCalculation(calculationName="Available Balance (Uncommitted)",
                       summariseExpression="sum(`Remaining Available Balance`)")
pt.6$evaluatePivot()


pt.risksummary <- PivotTable$new(argumentCheckMode='minimal')
pt.risksummary$addData(data)
pt.risksummary$addRowDataGroups("Disbursement Risk Level") 
pt.risksummary$defineCalculation(calculationName="Grant Count",
                       summariseExpression="n()")
pt.risksummary$defineCalculation(calculationName="Total Grant Amount",
                       summariseExpression="sum(`Grant Amount`)")
pt.risksummary$defineCalculation(calculationName="Total Available Balance Uncommitted",
                       summariseExpression="sum(`Remaining Available Balance`)",format= function(x) {dollar(x)})
pt.risksummary$defineCalculation(calculationName="Available Balance to Implement by 8/2020",
                       summariseExpression="sum(subset_available_bal)")
pt.risksummary$defineCalculation(calculationName="PO Commitments to spend by 8/2020",
                       summariseExpression="sum(subset_po_commitments)")
pt.risksummary$evaluatePivot()
pt.risksummary$renderPivot()





message("all pivots created")

wb <- createWorkbook()
addWorksheet(wb, "Summary")


start.1 <- 15
end.1 <- start.1 + (nrow(pt.1$asDataFrame())) + 1

start.2 <- end.1 + 6
end.2 <- start.2 + (nrow(pt.2$asDataFrame())) + 1

start.3 <- end.2 + 6
end.3 <- start.3 + (nrow(pt.3$asDataFrame())) + 1

left.1 <- 3

left.2 <- 8
writeDataTable(
  wb,
  sheet = "Summary",
  x = df.8,
  startCol = left.1,
  startRow = 4,
  withFilter = FALSE,
  rowNames = TRUE,
  tableName = 'test'
)

pt.risksummary$writeToExcelWorksheet(
  wb = wb,
  wsName = "Summary",
  topRowNumber = 4,
  leftMostColumnNumber = left.2,
  applyStyles =TRUE,outputValuesAs = "ACCOUNTING",
)

pt.1$writeToExcelWorksheet(
  wb= wb,
  wsName = "Summary",
  topRowNumber = start.1,
  leftMostColumnNumber = left.1-1,
  applyStyles = TRUE
)


pt.2$writeToExcelWorksheet(
  wb= wb,
  wsName = "Summary",
  topRowNumber = start.2,
  leftMostColumnNumber = left.1,
  applyStyles = TRUE
)


pt.3$writeToExcelWorksheet(
  wb= wb,
  wsName = "Summary",
  topRowNumber = start.3,
  leftMostColumnNumber = left.1,
  applyStyles = TRUE
)

pt.4$writeToExcelWorksheet(
  wb= wb,
  wsName = "Summary",
  topRowNumber = start.1,
  leftMostColumnNumber = left.2,
  applyStyles = TRUE
)


pt.5$writeToExcelWorksheet(
  wb= wb,
  wsName = "Summary",
  topRowNumber = start.2,
  leftMostColumnNumber = left.2+1,
  applyStyles = TRUE
)


pt.6$writeToExcelWorksheet(
  wb= wb,
  wsName = "Summary",
  topRowNumber = start.3,
  leftMostColumnNumber = left.2+1,
  applyStyles = TRUE
)


excel_df <- data %>% select(
  `Window #`,
  `Window Name`,
  `Trustee #`,
  `Trustee Fund Name`,
  `Child Fund #`,
  `Child Fund Name`,
  `Project ID`,
  `Execution Type`,
  `Child Fund Status`,
  `Child Fund TTL Name`,
  `TTL Unit`,
  `TTL Unit Name`,
  `Managing Unit`,
  `Lead GP/Global Theme`,
  `GPURL_binary`,
  `VPU`	,
  `Region Name`,
  `Country`	,
  `Activation Date`,
  `Closing Date`,
  `Grant Amount`,
  `Cumulative Disbursements`,
  `2020 Disbursements`,
  `PO Commitments`,
  `Remaining Available Balance`,
  `Transfer-in USD`	,
  `Transfers-out in USD`,
  `Real Transfer-in`,
  `Not Yet Transferred`,
  `Available Balance USD`,
  `Months to Closing Date`,
  `closed_or_not`
)

excel_df <- excel_df %>% rename("Available Balance (Uncommitted)" = `Remaining Available Balance`)

addWorksheet(wb, "MASTER grants")
writeDataTable(wb,"MASTER grants",excel_df,
               startRow = 2,
               startCol = 1,)

low_risk <- createStyle(fgFill ="#86F9B7",halign = "left")
medium_risk <- createStyle(fgFill ="#FFE285",halign = "left")
high_risk <- createStyle(fgFill ="#FD8D75",halign = "left")
very_high_risk <- createStyle(fgFill ="#F17979",halign = "left")
grace_period_style <- createStyle(fgFill ="#B7B7B7",halign = "left")


dollar_format <- createStyle(numFmt = "ACCOUNTING")
percent_format <- createStyle(numFmt = "0%")
date_format <- createStyle(numFmt = "mm/dd/yyyy")
num_format <- createStyle(numFmt = "NUMBER")
wrapped_text <- createStyle(wrapText = TRUE,valign = "top")
black_and_bold <- createStyle(fontColour = "#000000",textDecoration = 'bold')
merge_wrap <- createStyle(halign = 'center',valign = 'center',wrapText = TRUE)
gen_style <- createStyle(numFmt = "GENERAL")


addStyle(wb,1,rows = 5,cols = 8,style = very_high_risk,stack = T)
addStyle(wb,1,rows = 6,cols = 8,style = high_risk,stack = T)
addStyle(wb,1,rows = 7,cols = 8,style = medium_risk,stack = T)
addStyle(wb,1,rows = 8,cols = 8,style = low_risk,stack = T)
addStyle(wb,1,rows = 9,cols = 8,style = grace_period_style,stack = T)


addStyle(wb,1,rows = c(start.1,start.2,start.3),cols = 10,style = very_high_risk,stack = T)
addStyle(wb,1,rows = c(start.1,start.2,start.3),cols = 11,style = high_risk,stack = T)
addStyle(wb,1,rows = c(start.1,start.2,start.3),cols = 12,style = medium_risk,stack = T)
addStyle(wb,1,rows = c(start.1,start.2,start.3),cols = 13,style = low_risk,stack = T)
addStyle(wb,1,rows = c(start.1,start.2,start.3),cols = 14,style = grace_period_style,stack = T)



conditionally_format_rows <- function(x, start_row,colz){
  
  df <- x$asDataFrame()
  cond_rows <-
    which(!(str_trim(stri_replace_all_regex(row.names(df), ".[1-9]", ""),
                     side = "right") 
            %in% c("Total", unique(data$`Region Name`)))) + start_row
  
  
  var <- cond_rows[1]
  rangez[[1]] <- (cond_rows[1])
  temp_count <- 1
  for (i in cond_rows[-1]){
    
    if(i-1==var){
      rangez[[temp_count]] <- c(rangez[[temp_count]],i)
    }
    else{
      temp_count <- temp_count+1
      rangez[[temp_count]] <- i
    }
    var = i
  }
  
for (i in 1:length(rangez)) {

conditionalFormatting(wb, "Summary",
                      cols=colz,
                      rows=rangez[[i]],
                      style = c("white","red"),
                      type = "colourScale")
}

  message("Finished conditionally formatting this bad boy")
}

conditionally_format_rows(pt.1,start.1,6)
conditionally_format_rows(pt.2,start.2,6)
conditionally_format_rows(pt.2,start.3,6)
conditionally_format_rows(pt.4,start.1,10:11)
conditionally_format_rows(pt.5,start.2,10:11)
conditionally_format_rows(pt.6,start.3,10:11)


cols_r <-  c(5,6,10:15)
rows_r <- 6:end.3

for (i in cols_r){
  addStyle(wb,1,style = dollar_format,stack = T,rows = rows_r, cols = i)
}


addStyle(wb,1,style=dollar_format,stack=TRUE,rows=5,cols=10:15)
addStyle(wb,1,style=gen_style,stack=TRUE,rows=5,cols=4:6)
addStyle(wb,1,style = dollar_format,stack = T,rows = 6:9, cols = 4)
addStyle(wb,1,style = percent_format,stack = T,rows = 8, cols = 4:6)




setColWidths(wb,1,9:13,widths = 15 )
addStyle(wb,1,wrapped_text,4,9:13,stack = T)
setColWidths(wb,1,4:15,widths = 'auto')
setColWidths(wb,1,3,widths = 20 )
addStyle(wb,1,wrapped_text,5:9,3,stack = T)
#deleteData(wb,1,3,4)






