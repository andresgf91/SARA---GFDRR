# n_months <- PMA_grants$months_to_end_disbursement_static[which.max(PMA_grants$months_to_end_disbursement_static)]
# list_monthly_resources <- rep(list(0),n_months)
# list_grant_df <- rep(list(0),nrow(PMA_grants))
# first_date <- as_date(date_data_udpated) %>% floor_date(unit = "months")

# for (i in 1:nrow(PMA_grants)){
#   
#   print(i)
#   temp_max_months <- PMA_grants$months_to_end_disbursement_static[i]
#   temp_max_months <- ifelse(temp_max_months<=0,1,temp_max_months)
#   monthly_allocation <- (PMA_grants$unnacounted_amount[i])/(ifelse(temp_max_months>0,
#                                                                    temp_max_months,
#                                                                    1))
#   
#   temp_df <- data_frame(fund=rep(as.character(PMA_grants$Fund[i]),temp_max_months),
#                         fund_name=rep(as.character(PMA_grants$`Fund Name`[i]),temp_max_months),
#                         trustee=rep(as.character(PMA_grants$Trustee[i]),temp_max_months),
#                         fund_TTL=rep(as.character(PMA_grants$`Fund TTL Name`[i]),temp_max_months),
#                         sub_date = first_date %m+% months(c(0:(temp_max_months-1))),
#                         amount=rep(monthly_allocation,temp_max_months)
#   )
#   
#   list_grant_df[[i]] <- temp_df
#   
#   for (j in 1:temp_max_months){
#     list_monthly_resources[[j]] <- list_monthly_resources[[j]] + monthly_allocation
#   }
#   
#   #print(list_monthly_resources[[1]])
# }

#gg_data <- do.call(rbind,list_grant_df)
# 
# data <- data.frame(month_num=1:n_months,
#                    amount_available=unlist(list_monthly_resources))
# 
# data$dates <- first_date %m+% months(c(0:(nrow(data)-1)))
# 
# data$yearmonth <- paste0(year(data$dates),ifelse(month(data$dates)<10,
#                                                  paste0(0,month(data$dates)),
#                                                  month(data$dates)))
# 
# gg_data$yearmonth <- paste0(year(gg_data$sub_date),ifelse(month(gg_data$sub_date)<10,
#                                                           paste0(0,month(gg_data$sub_date)),
#                                                           month(gg_data$sub_date)))
# 
# data$yearquarter<- paste0(as.numeric(year(data$dates)),as.numeric(quarter(data$dates)))
# gg_data$yearquarter<- paste0(as.numeric(year(gg_data$sub_date)),
#                              as.numeric(quarter(gg_data$sub_date)))
# 
# data$quarterr <- zoo::as.yearqtr(data$dates)
# gg_data$quarterr <- zoo::as.yearqtr(gg_data$sub_date)
# 
# gg_data$quarter_lubri <- lubridate::quarter(gg_data$sub_date,with_year = TRUE)
# 
# grouped_gg_data <- gg_data %>%
#   mutate(quarterr=as_date(quarterr)) %>% 
#   group_by(quarterr,fund_name,fund,yearquarter) %>%
#   summarise(amount=sum(amount))
# 
# 
# quarterly_total <- gg_data %>% mutate(quarterr=as_date(quarterr)) %>%
#   group_by(quarterr) %>% summarise(Q_amount=sum(amount))
# 
# grouped_gg_data <- full_join(grouped_gg_data,quarterly_total,by="quarterr")
# 
# gg_df <- grouped_gg_data %>% filter(yearquarter<20210)