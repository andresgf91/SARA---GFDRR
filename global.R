# PACKAGES ----------------------------------------------------------------
library(readxl)
library(dplyr)
library(stringi)
library(lubridate)
library(pivottabler)
library(stringr)
library(googlesheets4)
library(googledrive)

options(scipen = 999)
# READ IN DATA ------------------------------------------------------------

#NEEDS UPDATING

drive_auth(use_oob = TRUE,
           path="sara-gfdrr-71de4c83b517.json")

SAP_folder_link <- "https://drive.google.com/open?id=16sUqxA8FzTS5U0LQTXSGEIMF9Ejo8niw"

#get list of folders (dates SAP updated) by 
docs_df <- drive_ls(SAP_folder_link)[,1:2] %>%
  as.data.frame()


docs_df$date <- docs_df$name %>% mdy()

docs_df <- docs_df %>%
  dplyr::arrange(desc(date))

latest_files_df <-  drive_ls(as_id(docs_df$id[1])) %>%
  select(name,id)

previous_files_df <-  drive_ls(as_id(docs_df$id[3])) %>%
  select(name,id)

grant_id <- which(stri_detect_fixed(tolower(latest_files_df$name),pattern = "grant")) %>%
  latest_files_df$id[.] %>% as_id()

trustee_id <- which(stri_detect_fixed(tolower(latest_files_df$name),pattern = "trustee")) %>%
  latest_files_df$id[.] %>% as_id()

previous_grant_id <- which(stri_detect_fixed(tolower(previous_files_df$name),pattern = "grant")) %>%
  latest_files_df$id[.] %>% as_id()

previous_trustee_id <- which(stri_detect_fixed(tolower(previous_files_df$name),pattern = "trustee")) %>%
  latest_files_df$id[.] %>% as_id()


drive_download(file = grant_id,
               path="data/latest_grant_data.xlsx",
               overwrite = TRUE)


drive_download(file = trustee_id,
               path="data/latest_trustee_data.xlsx",
               overwrite = TRUE)


drive_download(file = previous_grant_id,
               path="data/previous_grant_data.xlsx",
               overwrite = TRUE)


drive_download(file = previous_trustee_id,
               path="data/previous_trustee_data.xlsx",
               overwrite = TRUE)

old_grants_file <- "data/previous_grant_data.xlsx"
old_trustee_file <- "data/previous_trustee_data.xlsx"

grants_file <- "data/latest_grant_data.xlsx"
trustee_file <- "data/latest_trustee_data.xlsx"

#grants_file <- "data/GFDRR_test Sara Grant Level Data 4_27_20.xlsx"

JAIME_preprocess <- FALSE

fp_raw_data <- read_xlsx(path = "data/FP_GFDRR Raw Data 3_4_2020.xlsx",sheet = 2,skip = 6)
recode_trustee <- read_xlsx('data/recodes.xlsx',sheet=1)
recode_region <- read_xlsx('data/recodes.xlsx',sheet=2)
recode_GT <- read_xlsx("data/Global Theme - Resp. Unit Mapping.xlsx")


# glossary <- read_sheet("https://docs.google.com/spreadsheets/d/1S9Rpd4iIvrTvgyvCXYcRG_7BaFV82mPYo_557yVfSO4/edit?usp=sharing",sheet = 1)
# 
# write.csv(glossary,"data/glossary_1.csv")

gloss_banners <- read_xlsx("data/Glossary_SARA.xlsx",sheet = 'Tab Banners')
gloss_terms <- read_xlsx("data/Glossary_SARA.xlsx",sheet = 'Terminology') %>% arrange(Term)


##GLOBAL FUNCTIONS ---------

remove_periods_from_names <- function(df){
  names_with_dots <- sum(stri_detect_fixed(names(df),"."))
  
  if (names_with_dots > 0) {
    names(df) <- stri_replace_all_fixed(names(df),"."," ")
  }
  
  df
}

all_fun <- function(INPUT,CHOICES){
  if (INPUT == "ALL"){
    
    return(CHOICES)}
  else{
    
    return(INPUT)}
  
}


#Data processing -----

process_data_files <- function(GRANTS_file = grants_file,TRUSTEE_file = trustee_file,date_pos =1) {
  
  print(GRANTS_file)

grants <- read_xlsx(GRANTS_file)
trustee <- read_xlsx(TRUSTEE_file)



#date_data_udpated <- lubridate::mdy(stri_sub(GRANTS_file,from = -15,-6))
date_data_udpated <- docs_df$date[date_pos]

report_data_date <- paste0(
  month(date_data_udpated, abbr = FALSE, label = TRUE),
  " ",
  day(date_data_udpated),
  ", ",
  year(date_data_udpated)
)


grants <- remove_periods_from_names(grants)


swedish_grants_to_remove <- c("TF070808","TF080129")


#rename_grants in GRANTS 

grant_names <- names(grants)

print(grant_names)

if ("Lead GP/Global Theme" %in% grant_names){
  grants <- grants %>% dplyr::rename("Lead GP/Global Themes"=`Lead GP/Global Theme`)
  message("Changed col name LEAD GP/Global Theme")
}

if ("Child Fund" %in% grant_names){
  grants <- grants %>% dplyr::rename("Fund"=`Child Fund`)
  message("Changed col name FUND")
}

if ("Child Fund Name" %in% grant_names){
  grants <- grants %>% dplyr::rename("Fund Name"=`Child Fund Name`)
  message("Changed col name CHILD FUND NAME")
}

if ("Child Fund Status" %in% grant_names){
  grants <- grants %>% dplyr::rename("Fund Status"=`Child Fund Status`)
  message("Changed col name FUND STATUS")
}

if ("Child Fund TTL Name" %in% grant_names){
  grants <- grants %>% dplyr::rename("Fund TTL Name"=`Child Fund TTL Name`)
  message("Changed col name FUND TTL NAME")
}

if ("Grant Amount" %in% grant_names){
  grants <- grants %>% dplyr::rename("Grant Amount USD"=`Grant Amount`)
  message("Changed col name GRANT AMOUT USD")
}

if ("Cumulative Disbursements" %in% grant_names){
  grants <- grants %>% dplyr::rename("Disbursements USD"=`Cumulative Disbursements`)
  message("Changed col name DISBURSEMENTS USD")
}

if ("PO Commitments" %in% grant_names){
  grants <- grants %>% dplyr::rename("Commitments USD"=`PO Commitments`)
  message("Changed col name Commitments USD")
}

if ("Execution Type" %in% grant_names){
  grants <- grants %>% dplyr::rename("DF Execution Type"=`Execution Type`)
  message("Changed col name DF Execution Type")
}

if ("2020 Disbursements" %in% grant_names){
  grants <- grants %>% dplyr::rename("2020 Disbursement USD"=`2020 Disbursements`)
  message("Changed col name 2020 Disbursement USD")
}

if ("Transfer-in" %in% grant_names){
  grants$`Transfer-in USD` <- grants$`Transfer-in`
  message("ADDED col TRANSFER IN USD")
}

if ("Transfers-out" %in% grant_names){
  grants$`Transfers-out in USD` <-  grants$`Transfers-out`
  message("ADDED col Transfers-out in USD")
}

if (!("Fund Managing Unit Name" %in% grant_names)){
  
  grants$`Fund Managing Unit Name` <- grants$`TTL Unit Name`

}

if (!("Child Fund Managing Unit" %in% grant_names)){
  
  grants$`Child Fund Managing Unit` <- 
    ifelse("Child Fund Managing Unit Name" %in% grant_names,
           grants$`Child Fund Managing Unit Name`,
           ifelse("Fund Managing Unit Name" %in% grant_names,
                  grants$`Fund Managing Unit Name`, 
           ifelse("TTL Unit Name" %in% grant_names,
                  grants$`TTL Unit Name`,
                  "No managing unit in the data")
           )
  )
}

print(names(grants))

#rename columns in trustee df if neccessary
trustee_names <- names(trustee)
if ("Contribution Agreement Signed (Ledger) U...19" %in% trustee_names){
  trustee <- trustee%>% 
    dplyr::rename("Net Signed Condribution in USD"=`Contribution Agreement Signed (Ledger) U...19`)
}
if ("Contribution Cash received (ledger) USD" %in% trustee_names){
  trustee <- trustee %>%
    dplyr::rename("Net Paid-In Condribution in USD"=`Contribution Cash received (ledger) USD`)
}
if ("Contribution Agreement Signed (Ledger) U...21" %in% trustee_names){
  trustee <- trustee %>%
    dplyr::rename("Net Unpaid contribution in USD"=`Contribution Agreement Signed (Ledger) U...21`)
}


grants <- grants %>% filter(!(Fund %in% swedish_grants_to_remove))

#ADDING LEAD GP/GLOBAL THEMES-----------

# Change name of CLimate Change to Climate Change/GFDRR
recode_GT$`Lead GP/Global Themes`[which(recode_GT$`Lead GP/Global Themes`=="Climate Change")] <- "Climate Change/GFDRR" 
grants$`Lead GP/Global Themes`[which(grants$`Lead GP/Global Themes`=="Climate Change")] <- "Climate Change/GFDRR"


#create a new df with only active trustees and with recoded names
trustee$still_days_to_disburse <- as.Date(trustee$`TF End Disb Date`) > as.Date(date_data_udpated)
active_trustee <- trustee %>% filter(still_days_to_disburse == TRUE)

active_trustee <- left_join(active_trustee,recode_trustee,by=c("Fund"="Trustee")) %>%
  rename("Trustee.name"=`Trustee Fund Name`)

#add short region name to SAP data 


if("Fund Country Region Name" %in% grant_names){
  
grants <- full_join(grants,recode_region,by=c("Fund Country Region Name"='Region_Name'))}

#filter out grants that have 0 or that are not considered "Active" as per GFDRR definition which includes PEND

grants$`Available Balance USD` <- grants$`Grant Amount USD` -
  (grants$`Commitments USD` + grants$`Disbursements USD`)



raw_grants <- grants

grants <- grants %>% filter(`Fund Status` %in% c("ACTV","PEND"))

#create a  new column with pseudo final TRUSTEE names to display
#this is just in case not all desired names have been provided by GFDRR

active_trustee$temp.name <- NA
for (i in 1:nrow(active_trustee)){
  if (is.na(active_trustee$Trustee.name[i])){
    active_trustee$temp.name[i] <- active_trustee$`Fund Name`[i]
  } else
  {active_trustee$temp.name[i] <- active_trustee$Trustee.name[i]}
}



elapsed_months <- function(end_date, start_date) {
  ed <- as.POSIXlt(end_date)
  sd <- as.POSIXlt(start_date)
  12 * (ed$year - sd$year) + (ed$mon - sd$mon)
}


active_trustee$months_to_end_disbursement_dynamic <- 
  elapsed_months(active_trustee$`TF End Disb Date`,today())

active_trustee$months_to_end_disbursement_static <- as.numeric(floor(
  (difftime(
    strptime(active_trustee$`TF End Disb Date`, format = "%Y-%m-%d",tz = "GMT"),
    strptime(date_data_udpated, format = "%Y-%m-%d",tz="GMT"),
    units="days")
  )/(365.25/12)))



active_trustee$`Donor Agency Name` <- tools::toTitleCase(tolower(active_trustee$`Donor Agency Name`))

active_trustee$`Donor Agency Name` <- ifelse(active_trustee$`Donor Agency Name`=='Multi Donor',
                                             "Multiple Donors",
                                             active_trustee$`Donor Agency Name`)

grants$months_to_closing_dynamic <- 
  elapsed_months(grants$`Closing Date`,today())

grants$`Months to Closing Date`<- as.numeric(floor(
  (difftime(
    strptime(grants$`Closing Date`, format = "%Y-%m-%d",tz = "GMT"),
    strptime(date_data_udpated, format = "%Y-%m-%d",tz="GMT"),
    units="days")
  )/(365.25/12)))


grants$`Months to Closing Date` <- ifelse(grants$`Months to Closing Date`== 0,
                                                 1,
                                                 grants$`Months to Closing Date`)


grants$closed_or_not <- ifelse(grants$`Closing Date`< date_data_udpated,"yes","no") 

# add a column with Trustee temporary name V
grants <- left_join(grants,active_trustee %>%
                   filter(Fund %in% trustee$Fund) %>% 
                   select(temp.name,Fund),
                 by = c("Trustee"="Fund")) %>%
  filter(!is.na(Fund))


if(length(which(is.na(grants$temp.name)))>0){
  nameless <- unique(grants$Trustee[which(is.na(grants$temp.name))])
  
  for (i in nameless){
    
    if(i %in% recode_trustee$Trustee){
    grants$temp.name[grants$Trustee==i] <- recode_trustee$`Trustee Fund Name`[recode_trustee$Trustee==i]
      }
    }
  
}

grants$closing_FY <- ifelse(month(grants$`Closing Date`) > 6,
                         year(grants$`Closing Date`) + 1,
                         year(grants$`Closing Date`))

grants$activation_FY <- ifelse(month(grants$`Activation Date`) > 6,
                            year(grants$`Activation Date`) + 1,
                            year(grants$`Activation Date`))

grants$remaining_balance <- grants$`Grant Amount USD` - (grants$`Disbursements USD` + grants$`Commitments USD`)

grants$unnacounted_amount <- grants$`Grant Amount USD` - (grants$`Disbursements USD` + grants$`Commitments USD`)

grants$percent_unaccounted <- (grants$unnacounted_amount * 100)/grants$`Grant Amount USD`

grants$tf_age_months <- elapsed_months(date_data_udpated,grants$`Activation Date`)

grants$planned_tf_duration_months <- elapsed_months(grants$`Closing Date`,grants$`Activation Date`)

grants$time_ratio <- (grants$tf_age_months/grants$planned_tf_duration_months)

grants$disbursement_rate <- ifelse(grants$`Disbursements USD`>0,
                                           (grants$`Disbursements USD`/grants$`Grant Amount USD`),
                                           0)
grants$monthly_disbursement_rate <- ifelse(grants$`Disbursements USD`>0,
                                   (grants$`Disbursements USD`/grants$`Grant Amount USD`)/grants$tf_age_months,
                                   0)

grants$completion_gap <- grants$time_ratio - grants$disbursement_rate

grants$burn_rate <- (grants$`Disbursements USD`/ grants$tf_age_months)

grants$percent_remaining <- grants$remaining_balance/grants$`Grant Amount USD`

grants$required_disbursement_rate <- (grants$unnacounted_amount/grants$`Grant Amount USD`)/
  (ifelse(grants$`Months to Closing Date`>1,grants$`Months to Closing Date`,1))

grants$required_disbursement_rate <- ifelse(grants$required_disbursement_rate==Inf,
                                      grants$percent_unaccounted/100,
                                      grants$required_disbursement_rate)

grants$required_disbursement_rate <- ifelse(grants$closed_or_not == "yes",
                                            999,
                                            grants$required_disbursement_rate)


compute_risk_level <- function (x){
risk_level <- ifelse(x == 999, "Closed (Grace Period)",
                     ifelse(x < .025,"Low Risk",
                            ifelse(x < .055,"Medium Risk",
                                   ifelse(x < .105,"High Risk",
                                          "Very High Risk"))))

return(risk_level)
}

grants$disbursement_risk_level <- compute_risk_level(grants$required_disbursement_rate)

grants$disbursement_risk_level[is.na(grants$disbursement_risk_level)] <- "Low Risk"

compute_risk_color<- function (x){
  
  risk_color <- ifelse(x==999, "azure1",
                       ifelse(x<.025,"green",
                       ifelse(x<.055,"yellow",
                              ifelse(x<.105,"orange",
                                     "red"))))
  
  return(risk_color)
}

grants$disbursement_risk_color <- compute_risk_color(grants$required_disbursement_rate) 

#compute adjusted transfers (accounting for transfers out)

grants$`Real Transfers in` <- grants$`Transfer-in USD` - grants$`Transfers-out in USD`

#compute percentage of what was transferred that is available
grants$percent_transferred_available <- grants$unnacounted_amount/grants$`Real Transfers in`

#amount that is still available to be transferred
grants$`Not Yet Transferred` <- grants$`Grant Amount USD` - grants$`Real Transfers in`

#compute percentage of total grant that still needs to be transferred
grants$percent_left_to_transfer <- grants$`Not Yet Transferred`/grants$`Grant Amount USD`

#two step process to identify PMA grants (OR JUST ONE IF YOU WANT TO KEEP JUST-IN-TIME grants in):

#step 1
grants$PMA <- ifelse(is.na(grants$`Project ID`),'yes','no')

#this is the just-in-time grant that need to be 
grants$PMA[grants$Fund=="TF018938"] <- 'no' 

#step 2
#remove_just-in-time from PMA 
#grants$PMA <- ifelse(stringi::stri_detect(tolower(grants$`Fund Name`),
                   ##                       fixed=tolower('Just-in-Time')),
                   #  'yes',
                   #  grants$PMA)

PMA_grants <- grants %>% filter(PMA=='yes',`Grant Amount USD`>0)


grants$PMA.2 <- ifelse(grants$PMA=="yes","PMA","Operational")

grants <- grants %>% mutate(GPURL_binary = ifelse(`Lead GP/Global Themes`=="Urban, Resilience and Land",
                               "GPURL",
                               "Non-GPURL"))

grants$GPURL_binary[is.na(grants$GPURL_binary)] <- "Non-GPURL"
 
#set REGION colors as per GFDRR's requested scheme
regions_col_df <- data.frame(
  Region = c("LCR",
             "AFR",
             "GLOBAL",
             "EAP",
             "SAR",
             "ECA",
             "MNA"),
  region_color = c(
    "#f8a65b",
    "#f37960",
    "#adbac1",
    "#3487aa",
    "#57b6d0",
    "#96b263",
    "#be8cb0"
  ),
  stringsAsFactors = F
)


grants <- left_join(grants,regions_col_df,by="Region")
grants$`Remaining Available Balance` <- grants$unnacounted_amount

## REPORT PRODUCTION DATA PRE-PROCESSING 
report_grants <- grants %>%
  rename(
    "Child Fund #" = Fund,
    "Child Fund Name" = `Fund Name`,
    "Child Fund Status" = `Fund Status`,
    "Child Fund TTL Name" = `Fund TTL Name`,
    "Managing Unit" = `Child Fund Managing Unit`,
    "Closing FY" = closing_FY,
    "Region Name" = Region,
    "Trustee Fund Name" = temp.name,
    "Grant Amount" = `Grant Amount USD`,
    "Cumulative Disbursements" = `Disbursements USD`,
    "PO Commitments" = `Commitments USD`,
    "Execution Type" = `DF Execution Type`,
    "Window #" = `Window Number`,
    "Trustee #" = Trustee,
    "Lead GP/Global Theme" = `Lead GP/Global Themes`,
    "2020 Disbursements" = `2020 Disbursement USD`,
    "Required Monthly Disbursement Rate" = required_disbursement_rate,
    "Disbursement Risk Level" = disbursement_risk_level
  ) %>% 
  mutate(`Window #` = as.numeric(`Window #`)
  )

r_names <- names(report_grants)

if  ("TTL Unit" %in% r_names) {
  
  report_grants <- report_grants %>% 
  mutate(`TTL Unit` = as.character(`TTL Unit`))
}



report_grants$`Disbursement Risk Level` <- factor(report_grants$`Disbursement Risk Level` ,
                                         levels = c( "Very High Risk",
                                                     "High Risk",
                                                     "Medium Risk",
                                                     "Low Risk",
                                                     "Closed (Grace Period)"))


#report_grants <- report_grants %>% filter(`Child Fund #`!="TF0B2168")

grants <- grants %>% filter(`Grant Amount USD` > 0)




processed_data <- list("raw_grants" = raw_grants,
                       "grants" = grants,
                       "report_grants" = report_grants,
                       "trustee"=trustee,
                       "active_trustee"=active_trustee,
                       "PMA_grants" = PMA_grants,
                       "date_data_udpated" = date_data_udpated,
                       "report_data_date" = report_data_date)

return(processed_data)

}

processed_data <- process_data_files()

raw_grants <- processed_data[['raw_grants']]
grants <- processed_data[['grants']]
report_grants <- processed_data[['report_grants']]
trustee <- processed_data[['trustee']]
active_trustee <- processed_data[['active_trustee']]
PMA_grants <- processed_data[['PMA_grants']]
date_data_udpated <- processed_data[['date_data_udpated']]
report_data_date <- processed_data[['report_data_date']]

risk_colors <- tibble(
"Very High Risk" = "#C70039",
"High Risk" = "#FF5733",
"Medium Risk" = "#FFC300",
"Low Risk" = "#86F9B7",
"Closed (Grace Period)" = "#B7B7B7")

current_FY <- ifelse(month(date_data_udpated) > 6,
                     year(date_data_udpated) + 1,
                     year(date_data_udpated))

rm(processed_data)


