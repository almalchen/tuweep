#This R script performs the data cleaning and summarizing for client's ONTL, APD, SFWC, OA, and CRS lines of insurance. The inputs are templates populated with the client's raw data. The outputs are intended to replicate exactly like the edited Excel files containing the loss runs and pivot tables. These outputs can be linked directly to the master file.

# Output files include: 
# - check for missing claims 
# - checks for inc and paid 
# - pivot tables for each LOB 
# - large loss summaries
# - comparison to prior on claim level 
# - comparison of pivot table to prior pivot table 



# Instructions: 
# - create appropriate directories 
# - Populate the template files with raw data 
# - set current_eval and prior_eval Section 2 
# - set appropriate file paths to current and prior files in Section 2
# - run this script



# Table of contents: 
# Section 1. load libraries
# Section 2. import raw data
# Section 3. process ONTL APD losses 
# Section 4. output ONTL APD summaries 
# Section 5. process OA SFWC losses 
# Section 6. output OA SFWC summaries





######################   Note: 

# This is the anonymized version of the real script used in production. The client name has been changed to Tuweep and directories and file paths have been modified for confidentiality reasons. The names of subsidiaries have been modified to "Company 1" and "Company 2".


################################# 1. Load Libraries ########################################

library(openxlsx) # For reading/writing xlsx files
library(readxl)
library(dplyr)
library(fuzzyjoin)
library(lubridate) # for handling dates
library(kutils)
library(tidyverse)
options(scipen = 999) # Turn off scientific notation





################################ 2. Import raw data ######################################
#set evaluation dates for current study 
current_eval <- as.Date("2021-08-31")
prior_eval <- as.Date("2021-05-31")

#current files 
source_path <- "U:/ARC-Analytic/PERSONAL/Alma Chen/Tupweep R/20210831/Input/"
output_path <- "U:/ARC-Analytic/PERSONAL/Alma Chen/Tuweep R/20210831/Output/"

OA_SFWC_input_file <- "20210831 Tuweep OA SFWC Losses Template.xlsx"
ONTL_APD_input_file <- "20210831 Tuweep ONTL APD Losses Template.xlsx"
policy_period_input_file <- "Tuweep R policy period table.xlsx"

#output files 
OA_SFWC_results_file <- "20210831 Cleaned Tuweep OA SFWC Data.xlsx"
ONTL_APD_results_file <- "20210831 Cleaned Tuweep ONTL APD Data.xlsx"


#prior files
prior_source_path <- "U:/ARC-Analytic/PERSONAL/Alma Chen/Tuweep R/20210531/Output/"
prior_OA_SFWC_data <-"20210531 Cleaned Tuweep OA SFWC Data.xlsx"
prior_ONTL_APD_data <- "20210531 Cleaned Tuweep ONTL APD Data.xlsx"



# Load data 
OA_SFWC_input <- openxlsx::read.xlsx(xlsxFile = paste0(source_path, OA_SFWC_input_file), 
                                     startRow = 1, colNames = TRUE, detectDates = TRUE)

ONTL_APD_input <- openxlsx::read.xlsx(xlsxFile = paste0(source_path, ONTL_APD_input_file), 
                                      startRow = 1, colNames = TRUE, detectDates = TRUE)

policy_period_input <- openxlsx::read.xlsx(xlsxFile = paste0(source_path, policy_period_input_file), 
                                           colNames = TRUE, detectDates = TRUE, sheet = 'Table')
lookup_PA_claims <- readxl::read_excel(path = paste0(source_path, policy_period_input_file), 
                                       sheet = "Lookups", range = "D3:E11", col_names = TRUE, guess_max = 150)
lookup_CRS_claims <- readxl::read_excel(path = paste0(source_path, policy_period_input_file), 
                                        sheet = "Lookups", range = "G3:H35", col_names = TRUE, guess_max = 150)

lookup_manual_increases <- readxl::read_excel(path = paste0(source_path, policy_period_input_file), 
                                              sheet = "Lookups", range = "J3:R29", col_names = TRUE, guess_max = 150)


#load prior data

prior_OA_SFWC_database <- openxlsx::read.xlsx(xlsxFile = paste0(prior_source_path, prior_OA_SFWC_data), 
                                              startRow = 1, sheet = "Database",colNames = TRUE, detectDates = TRUE)
prior_OA_pivot <- readxl::read_excel(path = paste0(prior_source_path, prior_OA_SFWC_data), 
                                     range = "A5:E21", sheet = "OA",col_names = TRUE, guess_max = 150)
prior_SFWC_pivot <- readxl::read_excel(path = paste0(prior_source_path, prior_OA_SFWC_data), 
                                       range = "A5:F21", sheet = "SFWC",col_names = TRUE, guess_max = 150)
prior_OA_CRS_pivot <- readxl::read_excel(path = paste0(prior_source_path, prior_OA_SFWC_data), 
                                         range = "A5:E21", sheet = "OA CRS",col_names = TRUE, guess_max = 150)



prior_ONTL_APD <- openxlsx::read.xlsx(xlsxFile = paste0(prior_source_path, prior_ONTL_APD_data), 
                                      startRow = 1, colNames = TRUE, detectDates = TRUE)

prior_ONTL_APD_occ_table <- openxlsx::read.xlsx(xlsxFile = paste0(prior_source_path, prior_ONTL_APD_data), 
                                                startRow = 1, sheet = "Occ Table",colNames = TRUE, detectDates = TRUE)
prior_ONTL_APD_pivot <- readxl::read_excel(path = paste0(prior_source_path, prior_ONTL_APD_data), 
                                           range = "A5:E21", sheet = "Summary",col_names = TRUE, guess_max = 150)





############################ 3. Process ONTL APD Losses     ################################

#initial data check and format the dates 

ONTL_APD_input$inc_check <- ONTL_APD_input$inc_total >= ONTL_APD_input$paid_total

ONTL_APD_input$os_check <- round(ONTL_APD_input$os_res_total,0) == round(ONTL_APD_input$inc_total-ONTL_APD_input$paid_total, 0)

loss_date <- ONTL_APD_input$loss_date
if (!is.Date(loss_date)){
  ONTL_APD_input$loss_date <-  as.Date(loss_date, origin = "1899-12-30")
}


report_date <- ONTL_APD_input$report_date
if (!is.Date(report_date)){
  ONTL_APD_input$report_date <-  as.Date(report_date, origin = "1899-12-30")
}



# create the Occ Table tab 

#select relevant cols
ONTL_APD_subset <- ONTL_APD_input[,c('occ_num', 'loss_date', 'claim_num','coverage', 'inc_total', 'paid_total', 'inc_check')]


large_loss_threshold <- 100000
ONTL_APD_byocc <- ONTL_APD_subset %>% 
  filter(coverage != "Non-Trucking Liability") %>%
  group_by(occ_num, loss_date) %>%  
  summarize(
    sum_inc_total = sum(inc_total),
    sum_paid_total = sum(paid_total), 
    count_claim = n_distinct(claim_num)
  ) %>% 
  mutate(
    loss_date = loss_date,
    neg_inc_flag = ifelse(sum_inc_total >= sum_paid_total, 0, 1), 
    incurred = ifelse(neg_inc_flag==1, sum_paid_total, sum_inc_total),
    paid = sum_paid_total, 
    closed = ifelse(sum_paid_total== sum_inc_total, 1, 0), 
    closed_with_payment = ifelse((closed==1 & sum_inc_total >0 ), 1, 0), 
    open = 1-closed, 
    nz_count = ifelse(sum_inc_total==0, 0, 1), 
    large = ifelse(sum_inc_total > large_loss_threshold, 1, 0))


NTL_byocc <- ONTL_APD_subset %>% 
  filter(coverage == "Non-Trucking Liability") %>%
  group_by(occ_num, loss_date) %>%  
  summarize(
    sum_inc_total = sum(inc_total),
    sum_paid_total = sum(paid_total), 
    count_claim = n_distinct(claim_num)
  ) %>% 
  mutate(
    neg_inc_flag = ifelse(sum_inc_total >= sum_paid_total, 0, 1),
    incurred = ifelse(neg_inc_flag==1, sum_paid_total, sum_inc_total),
    paid = sum_paid_total, 
    closed = ifelse(sum_paid_total== sum_inc_total, 1, 0), 
    open = 1-closed)


#assign PY
ONTL_APD_byocc <- fuzzy_left_join(ONTL_APD_byocc
                                  , policy_period_input, by=c(loss_date = "policy_pd_begin", loss_date="policy_pd_end"), match_fun = list(`>=`, `<=`))


NTL_byocc <- fuzzy_left_join(NTL_byocc
                             , policy_period_input, by=c(loss_date = "policy_pd_begin", loss_date="policy_pd_end"), match_fun = list(`>=`, `<=`))



#create the Prior tab 
ONTL_APD_prior <- subset(prior_ONTL_APD_occ_table, select = c("occ_num", "loss_date", "incurred", "paid", "count_claim", "policy_pd_begin"))
ONTL_APD_prior <- ONTL_APD_prior %>%
  mutate(
    missing = ifelse(occ_num %in% ONTL_APD_byocc$occ_num, 0, 1)) 


missing_cc <-sum(ONTL_APD_prior$missing)
missing_inc<-sum(ONTL_APD_prior$incurred[which(ONTL_APD_prior$missing==1)])
missing_paid <-sum(ONTL_APD_prior$paid[which(ONTL_APD_prior$missing==1)])


#add the comparisons to prior to Occ Table tab 
ONTL_APD_byocc <- ONTL_APD_byocc %>%
  mutate(
    index_to_prior = sapply(occ_num, function(x) which(ONTL_APD_prior$occ_num==x)[1]),
    prior_inc = ifelse(!is.na(index_to_prior), ONTL_APD_prior$incurred[index_to_prior], 0),
    prior_paid = ifelse(!is.na(index_to_prior), ONTL_APD_prior$paid[index_to_prior], 0),
    chg_inc = incurred - prior_inc,
    chg_paid = paid - prior_paid, 
    large_change = ifelse( abs(chg_inc) > 25000 | abs(chg_paid) > 25000, 1,0)
  )





#create pivot table 
#all coverages excl NTL 

ONTL_APD_pivot <- ONTL_APD_byocc %>% group_by(policy_pd_begin) %>%
  summarize(net_incurred = sum(incurred), 
            net_paid = sum(paid), 
            nz_count = sum(nz_count), 
            open_count = sum(open))



#change from prior function
chg_from_prior_pivot <-function(current_pivot, prior_pivot){
  change_pivot <- cbind(current_pivot[,1],round(current_pivot[-c(1)] - prior_pivot[-c(1)], 2))
  return(change_pivot)
}



#change from prior pivot

ONTL_APD_chg <- chg_from_prior_pivot(ONTL_APD_pivot, prior_ONTL_APD_pivot)


#large losses 
ONTL_APD_large_pivot <- subset(ONTL_APD_byocc[ONTL_APD_byocc$large ==1,], select =c("occ_num","loss_date", "incurred", "paid"))

#NTL Pivot
NTL_pivot <- NTL_byocc %>% group_by(policy_pd_begin) %>%
  summarize(net_incurred = sum(incurred), 
            net_paid = sum(paid), 
            nz_count = sum(count_claim), 
            open_count = sum(open))



################################# 4. Output ONTL APD #############################################################

ONTL_APD_wb <- createWorkbook()


addWorksheet(ONTL_APD_wb, sheetName = "Prior")
writeData(ONTL_APD_wb, sheet = 1, x = ONTL_APD_prior, startCol = 1, startRow = 6)
writeData(ONTL_APD_wb, sheet = 1, x = "Sum of missing inc:", startCol = 10, startRow = 2)
writeData(ONTL_APD_wb, sheet = 1, x = missing_inc, startCol = 11, startRow = 2)
writeData(ONTL_APD_wb, sheet = 1, x = "Sum of missing paid:", startCol = 10, startRow = 3)
writeData(ONTL_APD_wb, sheet = 1, x = missing_paid, startCol = 11, startRow = 3)
writeData(ONTL_APD_wb, sheet = 1, x = "Sum of missing cc:", startCol = 10, startRow = 4)
writeData(ONTL_APD_wb, sheet = 1, x = missing_cc, startCol = 11, startRow = 4)


addWorksheet(ONTL_APD_wb, sheetName = "Database")
writeData(ONTL_APD_wb, sheet = 2, x = ONTL_APD_input, startCol = 1, startRow = )

addWorksheet(ONTL_APD_wb, sheetName = "Occ Table")
writeData(ONTL_APD_wb, sheet = 3, x = ONTL_APD_byocc, startCol = 1, startRow = 1)



addWorksheet(ONTL_APD_wb, sheetName = "Summary")
writeData(ONTL_APD_wb, sheet = 4, x = "Coverage: All Coverages Excluding NTL", startCol = 1, startRow = 1)
writeData(ONTL_APD_wb, sheet = 4, x = paste("Current Eval: ", format(current_eval)), startCol = 1, startRow = 4)
writeData(ONTL_APD_wb, sheet = 4, x = ONTL_APD_pivot, startCol = 1, startRow = 5)
writeData(ONTL_APD_wb, sheet = 4, x = ONTL_APD_large_pivot, startCol = 1, startRow = 34)
writeData(ONTL_APD_wb, sheet = 4, x = paste("Prior Eval: ", format(prior_eval)), startCol = 7, startRow = 4)
writeData(ONTL_APD_wb, sheet = 4, x = prior_ONTL_APD_pivot, startCol = 7, startRow = 5)
writeData(ONTL_APD_wb, sheet = 4, x = "Change", startCol = 12, startRow = 4)
writeData(ONTL_APD_wb, sheet = 4, x = ONTL_APD_chg, startCol = 12, startRow = 5)
writeData(ONTL_APD_wb, sheet = 4, x = "Large Losses", startCol = 1, startRow = 32)
writeData(ONTL_APD_wb, sheet = 4, x = ONTL_APD_large_pivot, startCol = 1, startRow = 34)
writeData(ONTL_APD_wb, sheet = 4, x = "Coverage: NTL", startCol = 19, startRow = 1)
writeData(ONTL_APD_wb, sheet = 4, x = paste("Current Eval: ", format(current_eval)), startCol = 19, startRow = 4)
writeData(ONTL_APD_wb, sheet = 4, x = NTL_pivot, startCol = 19, startRow = 5)
options(xlsx.date.format="%Y-%m-%d")


saveWorkbook(ONTL_APD_wb, file = paste0(output_path, ONTL_APD_results_file))









######################################## 5. OA SFWC Losses ##############################################


OA_SFWC_large_threshold <- 150000

#covert acc date to date 

accident_date <- OA_SFWC_input$accident_date
if (!is.Date(loss_date)){
  OA_SFWC_input$accident_date <-  as.Date(accident_date, origin = "1899-12-30")
}


OA_SFWC_input = fuzzy_left_join(OA_SFWC_input, policy_period_input, by=c(accident_date = "policy_pd_begin", accident_date="policy_pd_end"), match_fun = list(`>=`, `<=`))
OA_SFWC_input$year = year(OA_SFWC_input$policy_pd_begin)
OA_SFWC_input<-subset(OA_SFWC_input, select = -c(policy_pd_begin, policy_pd_end))




#if lob = OA and claimant name is not in the list of PA claims, assign OA
#if lob = SFWC, assign SFWC 
# we are not longer doing PA
# I omitted check col because it depends on PA 


OA_SFWC_input <- OA_SFWC_input %>%
  mutate(
    OA = ifelse((lob=="OA" & !(location_level3_name %in% lookup_PA_claims$`Claimant Name`)), 1, 0), 
    SFWC = ifelse(lob=="SFWC", 1, 0),
    adj_company = ifelse(location_level3_name %in% lookup_CRS_claims$`Claimant Name`, "Company 2", location_level2_name), 
    nz_count = ifelse(total_paid_net + total_outstanding > 0, 1, 0), 
    open = ifelse(total_outstanding==0, 0, 1),
    closed = 1-open, 
    cwp = ifelse((total_outstanding==0 & total_paid_net > 0), 1, 0), 
    inc_check = ifelse(total_incurred >= total_paid_net, TRUE, FALSE), 
    os_check = ifelse(round(total_incurred - total_paid_net) == round(total_outstanding), TRUE, FALSE), 
    check_occ = sapply(claim_number, function(x) sum(x==claim_number)), 
    manual_adjustment = ifelse(claim_number %in% lookup_manual_increases$CLAIM_NUM, 1, 0),
    reserve_increase = ifelse(manual_adjustment == 0 | closed == 1, 0,replace_na(lookup_manual_increases$`Total`[match(claim_number, lookup_manual_increases$CLAIM_NUM)], 0)),
    incurred_w_increase = reserve_increase + total_incurred, 
    large = ifelse(total_incurred > OA_SFWC_large_threshold, 1, 0)
  )



# run this for sanity check
# > max(OA_SFWC_input$check_occ)
# [1] 1


#Add the Prior tab
OA_SFWC_prior <- subset(prior_OA_SFWC_database, select = c("claim_number", "total_paid_net", "total_incurred"))
OA_SFWC_prior <- OA_SFWC_prior %>%
  mutate(
    missing = ifelse(claim_number %in% OA_SFWC_input$claim_number, 0, 1)
  )



#add the comparisons to prior to Database tab 
OA_SFWC_input <- OA_SFWC_input %>%
  mutate(
    index_to_prior = sapply(claim_number, function(x) which(OA_SFWC_prior$claim_number==x)[1]),
    prior_inc = ifelse(!is.na(index_to_prior), OA_SFWC_prior$total_incurred[index_to_prior], 0),
    prior_paid = ifelse(!is.na(index_to_prior), OA_SFWC_prior$total_paid_net[index_to_prior], 0),
    chg_inc = total_incurred - prior_inc,
    chg_paid = total_paid_net - prior_paid, 
    large_change = ifelse( abs(chg_inc) > 50000 | abs(chg_paid) > 50000, 1,0)
  )





#SFWC pivot 

SFWC_pivot <- OA_SFWC_input %>% 
  filter(SFWC ==1) %>%
  group_by(year) %>%
  summarize(net_inc = sum(total_incurred), 
            net_paid = sum(total_paid_net), 
            nz_count = sum(nz_count),
            open = sum(open), 
            net_inc_w_increases = sum(incurred_w_increase))




#run this for sanity check
#colSums(SFWC_pivot)


SFWC_large_pivot <- subset(OA_SFWC_input[OA_SFWC_input$large==1 & OA_SFWC_input$SFWC ==1,], select = c("claim_number","accident_date", "total_incurred", "total_paid_net", "incurred_w_increase"))


SFWC_years <- seq(2010, 2021, 1) #select years to include in the state distribution 
SFWC_state_dist <- OA_SFWC_input %>%
  filter(SFWC ==1 & year %in% SFWC_years) %>%
  group_by(benefit_state) %>%
  summarize(net_inc_w_increases=sum(incurred_w_increase, na.rm = TRUE)) %>% 
  mutate(percentage = net_inc_w_increases/sum(net_inc_w_increases))





#OA pivot 

OA_pivot <- OA_SFWC_input %>% 
  filter(OA ==1 & adj_company == 'Company 1') %>%
  group_by(year) %>%
  summarize(net_inc = sum(total_incurred), 
            net_paid = sum(total_paid_net), 
            nz_count = sum(nz_count), 
            open = sum(open))

OA_large_pivot <- subset(OA_SFWC_input[OA_SFWC_input$large==1 & OA_SFWC_input$OA ==1,], select = c("claim_number","accident_date", "total_incurred", "total_paid_net"))


#OA CRS pivot 

OA_CRS_pivot <- OA_SFWC_input %>% 
  filter(OA ==1 & adj_company == 'Company 2') %>%
  group_by(year) %>%
  summarize(net_inc = sum(total_incurred), 
            net_paid = sum(total_paid_net), 
            nz_count = sum(nz_count), 
            open = sum(open))

OA_CRS_large_pivot <- subset(OA_SFWC_input[OA_SFWC_input$large==1 & OA_SFWC_input$OA ==1 & OA_SFWC_input$adj_company == "CENTRAL REFRIGERATED SERVICE",], select = c("claim_number","accident_date", "total_incurred", "total_paid_net"))




#chg from prior pivots
prior_OA_CRS_pivot <- prior_OA_CRS_pivot[complete.cases(prior_OA_CRS_pivot),] #remove NA rows

SFWC_chg <- chg_from_prior_pivot(SFWC_pivot, prior_SFWC_pivot)
OA_chg <- chg_from_prior_pivot(OA_pivot, prior_OA_pivot)
OA_CRS_chg <- chg_from_prior_pivot(OA_CRS_pivot, prior_OA_CRS_pivot) 











######################################### 6. OA SFWC output section #################################################

OA_SFWC_wb <- createWorkbook()

addWorksheet(OA_SFWC_wb, sheetName = "Prior")
writeData(OA_SFWC_wb, sheet = 1, x = OA_SFWC_prior, startCol = 1, startRow = 1)
writeData(OA_SFWC_wb, sheet = 1, x = paste("Num missing: ", sum(OA_SFWC_prior$missing)), startCol = 7, startRow = 1)

addWorksheet(OA_SFWC_wb, sheetName = "Database")
writeData(OA_SFWC_wb, sheet = 2, x = OA_SFWC_input, startCol = 1, startRow = 5)

addWorksheet(OA_SFWC_wb, sheetName = "SFWC")
writeData(OA_SFWC_wb, sheet = 3, x = "Coverage: SFWC", startCol = 1, startRow = 1)
writeData(OA_SFWC_wb, sheet = 3, x = paste("Current Eval: ", format(current_eval)), startCol = 1, startRow = 4)
writeData(OA_SFWC_wb, sheet = 3, x = SFWC_pivot, startCol = 1, startRow = 5)
writeData(OA_SFWC_wb, sheet = 3, x = paste("Prior Eval: ", format(prior_eval)), startCol = 9, startRow = 4)
writeData(OA_SFWC_wb, sheet = 3, x = prior_SFWC_pivot, startCol = 9, startRow = 5)
writeData(OA_SFWC_wb, sheet = 3, x = "Change", startCol = 16, startRow = 4)
writeData(OA_SFWC_wb, sheet = 3, x = SFWC_chg, startCol = 16, startRow = 5)
writeData(OA_SFWC_wb, sheet = 3, x = "Large Losses", startCol = 1, startRow = 32)
writeData(OA_SFWC_wb, sheet = 3, x = SFWC_large_pivot, startCol = 1, startRow = 34)
writeData(OA_SFWC_wb, sheet = 3, x = "State Distribution", startCol = 8, startRow = 32)
writeData(OA_SFWC_wb, sheet = 3, x = SFWC_state_dist, startCol = 8, startRow = 34)

addWorksheet(OA_SFWC_wb, sheetName = "OA")
writeData(OA_SFWC_wb, sheet = 4, x = "Coverage: OA", startCol = 1, startRow = 1)
writeData(OA_SFWC_wb, sheet = 4, x = "Company: Company 1", startCol = 1, startRow = 2)
writeData(OA_SFWC_wb, sheet = 4, x = paste("Current Eval: ", format(current_eval)), startCol = 1, startRow = 4)
writeData(OA_SFWC_wb, sheet = 4, x = OA_pivot, startCol = 1, startRow = 5)
writeData(OA_SFWC_wb, sheet = 4, x = paste("Prior Eval: ", format(prior_eval)), startCol = 9, startRow = 4)
writeData(OA_SFWC_wb, sheet = 4, x = prior_OA_pivot, startCol = 9, startRow = 5)
writeData(OA_SFWC_wb, sheet = 4, x = "Change", startCol = 16, startRow = 4)
writeData(OA_SFWC_wb, sheet = 4, x = OA_chg, startCol = 16, startRow = 5)
writeData(OA_SFWC_wb, sheet = 4, x = "Large Losses", startCol = 1, startRow = 32)
writeData(OA_SFWC_wb, sheet = 4, x = OA_large_pivot, startCol = 1, startRow = 34)


addWorksheet(OA_SFWC_wb, sheetName = "OA CRS")
writeData(OA_SFWC_wb, sheet = 5, x = "Coverage: OA CRS", startCol = 1, startRow = 1)
writeData(OA_SFWC_wb, sheet = 5, x = "Company: Company 2", startCol = 1, startRow = 2)
writeData(OA_SFWC_wb, sheet = 5, x = paste("Current Eval: ", format(current_eval)), startCol = 1, startRow = 4)
writeData(OA_SFWC_wb, sheet = 5, x = OA_CRS_pivot, startCol = 1, startRow = 5)
writeData(OA_SFWC_wb, sheet = 5, x = paste("Prior Eval: ", format(prior_eval)), startCol = 9, startRow = 4)
writeData(OA_SFWC_wb, sheet = 5, x = prior_OA_CRS_pivot, startCol = 9, startRow = 5)
writeData(OA_SFWC_wb, sheet = 5, x = "Change", startCol = 16, startRow = 4)
writeData(OA_SFWC_wb, sheet = 5, x = OA_CRS_chg, startCol = 16, startRow = 5)
writeData(OA_SFWC_wb, sheet = 5, x = "Large Losses", startCol = 1, startRow = 32)
writeData(OA_SFWC_wb, sheet = 5, x = OA_CRS_large_pivot, startCol = 1, startRow = 34)


saveWorkbook(OA_SFWC_wb, file = paste0(output_path, OA_SFWC_results_file))


