

# Analyzing Emergency Med data
# Last Updated 01_14_2022

rm(list=ls())



#Import Libraries
suppressMessages({
library(readxl)
library(writexl)
library(stringr)
library(tidyverse)
library(dplyr)
library(reshape2)
library(svDialogs)
library(formattable)
library(scales)
library(ggpubr)
library(knitr)
library(kableExtra)
library(rmarkdown)
})

# Work Directory
wrk.dir <- "J:/deans/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Pharmacy/Data/Epic Latest Reports/Daily Reports"
setwd(wrk.dir)


# Import emergency_med_groups data
emergency_med_groups <- read_excel(paste0(wrk.dir, "/Medication Classification 01-20-2022.xlsx"), col_names = TRUE, na = c("", "NA"))


# Current week date
dates_pattern <-  seq(Sys.Date()-10, Sys.Date(), by='day')

#Import the latest Med Inv REPO file 
inv_repo <- file.info(list.files(path = paste0(wrk.dir,"/REPO/Emergency_Inv_Repo"), full.names = T , 
                                 pattern =paste0("Covid Surge Meds Inventory Repo-", dates_pattern, collapse = "|"  )))
repo_file <- rownames(inv_repo)[which.max(inv_repo$ctime)]
inv_repo <- readRDS(repo_file)
inv_repo <- inv_repo %>% distinct()

# Check the most recent ReportDate
max(inv_repo$ReportDate)

# Change the date format
inv_repo <- inv_repo %>% mutate(ReportDate = as.Date(ReportDate))

# Import COVID Surge Med WINV data
inv_list <- file.info(list.files(path= paste0(wrk.dir,"/New Inventory"), full.names=TRUE, 
                                 pattern = paste0("COVID Surge Med WINV Balance_", dates_pattern, collapse = "|")))
inv_list <- inv_list[with(inv_list, order(as.POSIXct(ctime), decreasing = TRUE)), ]

dif_time <- difftime(max(inv_repo$ReportDate), Sys.Date()-1)
count <-  gsub('.*-([0-9]+).*','\\1', dif_time )

files <- rownames(inv_list)[1:count]
new_inventory_raw <- lapply(files, function(filename){read_excel(filename, col_names = TRUE, na = c("", "NA"))})

#Add Date Column
new_inventory_raw <- lapply(new_inventory_raw,transform, UpdateDate = as.Date(LAST_UPDATE_TIME,  format = "%d-%b-%y"))
new_inventory_raw <- lapply(new_inventory_raw,transform,ReportDate =  as.Date(max(LAST_UPDATE_TIME),origin = "2020-01-01")-1)


inv_daily_df <-  do.call(rbind.data.frame,  new_inventory_raw )
rm(inv_list, new_inventory_raw)

# Format inventory raw data ----------------------
# Create columns with the last date each location/inventory item was updated and the date of the report
# (Assumes at least one location/inventory item was updated on reporting date)
# Determine site of inventory location
inv_daily_df <- inv_daily_df %>% mutate(Site = ifelse(str_detect("MSHS COVID 19 Stockpile", INV_NAME), "MSHS Stockpile",
                                                      str_replace(str_extract(INV_NAME, "[A-Z]+(\\s|\\-)"), "\\s|\\-", "")),
                                        Site=ifelse(Site== "BI","MSBI" , Site))


## Remove BI Supply room 
inv_daily_df <- inv_daily_df %>% filter(INV_NAME != "BI Supply room")


# Extract inventory item concentration ---------------------
conc_pattern_1 <- "[0-9]+(\\.|\\,)*[0-9]*\\s(MCG/)[0-9]*(\\.|\\,)*[0-9]*\\s*(ML)"  # Pattern 1: x MCG/y ML
conc_pattern_2 <- "[0-9]+(\\.|\\,)*[0-9]*\\s(MG/)[0-9]*(\\.|\\,)*[0-9]*\\s*(ML)"  # Pattern 2: x MG/y ML
conc_pattern_3 <- "[0-9]+(\\.|\\,)*[0-9]*\\s(MG)\\s"  # Pattern 3: x MG 
conc_pattern_4 <- "[0-9]+(\\.|\\,)*[0-9]*\\s(UNIT/)[0-9]*(\\.|\\,)*[0-9]*\\s*(ML)"  # Pattern 4: x UNIT/ y ML


# Extract inventory concentration from PRD_NAME
inv_daily_df <- inv_daily_df %>% mutate(ConcExtract = ifelse(str_detect(PRD_NAME, conc_pattern_1), str_extract(PRD_NAME, conc_pattern_1),
                                                             ifelse(str_detect(PRD_NAME, conc_pattern_2), str_extract(PRD_NAME, conc_pattern_2),
                                                                    ifelse(str_detect(PRD_NAME, conc_pattern_3), str_extract(PRD_NAME, conc_pattern_3),
                                                                           ifelse(str_detect(PRD_NAME, conc_pattern_4), str_extract(PRD_NAME, conc_pattern_4), NA)))))




inv_daily_df <- inv_daily_df %>% mutate(ConcExtract= str_trim(ConcExtract))

# Fix concentrations with missing spaces or missing numeric values (ie, 200 MG/200ML, 10 MG/ML)
inv_daily_df <- inv_daily_df %>%  mutate(ConcExtract = ifelse(str_detect(ConcExtract, "([0-9])(ML)"), str_replace(ConcExtract, "([0-9])(ML)", "\\1 \\2"),
                                                              ifelse(str_detect(ConcExtract, "(/ML)"), str_replace(ConcExtract, "(/ML)", "/1 ML"), ConcExtract)))

# Split concentrations on spaces
inv_daily_df$ConcExt_split<- inv_daily_df$ConcExtract
inv_daily_df <- inv_daily_df %>% separate(ConcExt_split, c("ConcDoseSize", "b", "ConcVolUnit"), extra = "merge", fill = "right", sep = " ")
inv_daily_df <- inv_daily_df %>% separate(b, c( "ConcDoseUnit", "ConcVolSize"), extra = "merge", fill = "right", sep = "/")




# Format concentration columns into numerics
inv_daily_df <- inv_daily_df %>% mutate(ConcDoseSize = as.numeric(str_replace(ConcDoseSize, "\\,", "")),
                                        ConcVolSize =as.numeric(str_replace(ConcVolSize, "\\,", "")))

# Calculate normalize dose unit
inv_daily_df <- inv_daily_df %>% mutate(ConcDoseUnit=replace(ConcDoseUnit, is.na(ConcDoseUnit), ""),
                                        ConcVolUnit= replace(ConcVolUnit, is.na(ConcVolUnit),""))

inv_daily_df <- inv_daily_df %>% mutate(NormDoseSize = ifelse(is.na(ConcVolSize), ConcDoseSize, ConcDoseSize / ConcVolSize),
                                        NormDoseUnit = paste0(ConcDoseUnit,"/",ConcVolUnit))


inv_daily_df <- inv_daily_df %>% mutate(ConcDoseUnit=ifelse(ConcDoseUnit==  "", NA, ConcDoseUnit),
                                        ConcVolUnit= ifelse(ConcVolUnit== "", NA, ConcVolUnit ))




# Strip out inventory volume and sizes from Inventory Item Name
inv_size_pattern <- "\\,\\s[0-9]+(\\.|\\,)*[0-9]*\\s*(mL|g)"

# Extract inventory item size
inv_daily_df <- inv_daily_df %>% mutate(InvShortName = str_replace(PRD_NAME, inv_size_pattern, ""),
                                        InvSizeUnit = str_replace(str_extract(PRD_NAME, inv_size_pattern), "\\,\\s", ""))


# Replace NA with "1 Each" and fix inventory sizes with missing spaces (ie, 2mL) 
inv_daily_df <- inv_daily_df %>% mutate(InvSizeUnit = ifelse(is.na(InvSizeUnit), "1 Each", InvSizeUnit))

inv_daily_df <- inv_daily_df %>% mutate(InvSizeUnit = (str_replace(InvSizeUnit, "([0-9]+)(mL)", "\\1 \\2")))



# split InvSizeUnit into inventory item size and unit 
inv_daily_df$InvSizeUnit_split <- inv_daily_df$InvSizeUnit
inv_daily_df <- inv_daily_df  %>% separate(InvSizeUnit_split, c("InvSize", "InvUnit"), extra = "merge", fill = "right", sep = " ")

inv_daily_df <- inv_daily_df %>% mutate(InvSize=  as.numeric(as.character(InvSize)))


# Calculate total dose balance
inv_daily_df <- inv_daily_df %>% mutate(TotalDoseBalance = BALANCE * NormDoseSize * InvSize)



# Fix NDC ID and Code data alignment & NormDoseUnit = MG/ and /
inv_daily_df <- inv_daily_df %>% mutate(NDC_ID_ref = NDC_ID, NDC_CODE_ref = NDC_CODE) %>%
  mutate(NDC_ID = ifelse(str_detect(NDC_ID,"-"), NDC_CODE_ref, NDC_ID_ref), NDC_CODE = ifelse(str_detect(NDC_CODE,"-"), NDC_CODE_ref, NDC_ID_ref)) %>%
  mutate(NormDoseUnit = ifelse(NormDoseUnit == "/", NA, ifelse(NormDoseUnit == "MG/", "MG", NormDoseUnit)))


# Aggregate inventory data by site and NDC 
inv_site_summary <- inv_daily_df %>% group_by(ReportDate, Site, PRD_NAME, NDC_ID, NDC_CODE, ConcExtract, ConcDoseSize, ConcDoseUnit,
                                              ConcVolSize, ConcVolUnit, NormDoseSize, NormDoseUnit, InvShortName, InvSizeUnit, InvSize, InvUnit) %>%
  summarize(Balance = sum(BALANCE), TotalDoseBalance = sum(TotalDoseBalance)) %>% ungroup()


inv_site_summary$MedGroup <- toupper(gsub("([A-Za-z]+).*", "\\1", inv_site_summary$PRD_NAME))
inv_site_summary$MedGroup <- ifelse(inv_site_summary$MedGroup == "NOREPINEPHRINE", "NOREPINEPHRINE BITARTRATE", inv_site_summary$MedGroup)


# Import Inventory Balance Data
inv_site_summary <- inv_site_summary[!duplicated(inv_site_summary),]

inv_site_summary$med_class <- emergency_med_groups$Classification[match(inv_site_summary$MedGroup, emergency_med_groups$`Medication Group`)]
inv_site_summary <- inv_site_summary %>% filter(med_class == "EMERGENCY") %>% filter(ReportDate >= "2021-03-15")

#inv_site_summary <- inv_site_summary %>% filter(Site %in% c("MSB","MSBI","MSH","MSHS Stockpile","MSM","MSQ","MSW"))


# Bind today's data with repository
inv_final_repo <- rbind(inv_repo, inv_site_summary)
inv_final_repo <- inv_final_repo %>% distinct()


# Save the new repo
saveRDS(inv_final_repo, paste0(wrk.dir, "/REPO/Emergency_Inv_Repo\\Covid Surge Meds Inventory Repo-", Sys.Date(), ".RDS"))
#write_xlsx(inv_final_repo, path = paste0(wrk.dir, "\\REPO\\Emergency_Inv_Repo\\Covid Surge Meds Inventory Repo-", Sys.Date(), ".xlsx"))
rm(inv_repo, inv_site_summary, inv_daily_df)



#----------- Import and pre-process Med Admin data  --------------

#Import the latest REPO file 
med_repo <- file.info(list.files(path = paste0(wrk.dir,"/REPO/Emergency_Med_Repo"), full.names = T , pattern =paste0("Covid Surge Meds Admin Repo-", dates_pattern, collapse = "|"  )))
repo_file <- rownames(med_repo)[which.max(med_repo$ctime)]
med_repo <- readRDS(repo_file)
med_repo <- med_repo %>% distinct()


# check the Admin Date
max(med_repo$Admin_Date)

# Change Date format
med_repo <- med_repo %>% mutate(Admin_Date= as.Date(Admin_Date))

# Import COVID Surge Med Admin data
new_med_admin_list <- file.info(list.files(path= paste0(wrk.dir,"/New Med Admin"), full.names=TRUE, 
                                           pattern = paste0("COVID Surge Med Admin for Pharmacy Dashboard", dates_pattern, collapse = "|")))
new_med_admin_list <- new_med_admin_list[with(new_med_admin_list, order(as.POSIXct(ctime), decreasing = TRUE)), ]


files <- rownames(new_med_admin_list)[1:count]
new_med_admin_raw <- lapply(files, function(filename){
  read_excel(filename, col_names = TRUE, na = c("", "NA"))})

new_med_admin <-  do.call(rbind.data.frame, new_med_admin_raw )
rm(new_med_admin_list, new_med_admin_raw)


new_med_admin$MedGroup <- toupper(gsub("([A-Za-z]+).*", "\\1", new_med_admin$DISPINSABLE_MED_NAME))
new_med_admin$MedGroup <- ifelse(new_med_admin$MedGroup == "NOREPINEPHRINE",
                                 "NOREPINEPHRINE BITARTRATE", new_med_admin$MedGroup)

# Import COVID Surge Meds Administered Data
new_med_admin <- new_med_admin[!duplicated(new_med_admin), ]

new_med_admin$med_class <- emergency_med_groups$Classification[match(new_med_admin$MedGroup, emergency_med_groups$`Medication Group`)]
new_med_admin <- new_med_admin %>% filter(med_class == "EMERGENCY") %>% filter(as.Date(TAKEN_DATETIME, format="%d-%m-%y") >= "2021-03-15")


## Fill Missing NDC_ID
new_med_admin <- new_med_admin  %>% arrange(LOC_NAME, DISPINSABLE_MED_NAME)

new_med_admin <- new_med_admin  %>% group_by(LOC_NAME, DISPINSABLE_MED_NAME) %>% fill(NDC_ID,.direction = c("down", "up", "downup", "updown"))
new_med_admin <- new_med_admin  %>% group_by(LOC_NAME, DISPINSABLE_MED_NAME) %>% mutate(NDC_ID=ifelse(is.na(NDC_ID), NDC_ID[1], NDC_ID))
new_med_admin <- new_med_admin  %>% group_by(LOC_NAME, DISPINSABLE_MED_NAME) %>% mutate(NDC_ID=ifelse(is.na(NDC_ID), NDC_ID[n()], NDC_ID))

new_med_admin <- new_med_admin %>% ungroup()




# Data Exclusion Criteria: Exclude ADMIN_REASON = "Bolus from Infusion"
covid_meds_admin <- new_med_admin %>% filter(ADMIN_REASON != "Bolus from Infusion")

# Format columns
covid_meds_admin$Admin_Date <- as.Date(covid_meds_admin$TAKEN_DATETIME, format="%d-%m-%y")
covid_meds_admin$MAR_DOSE <- as.numeric(covid_meds_admin$MAR_DOSE)

# 2. Strip off the concentration from med name 
conc_pattern_1 <- "[0-9]+\\s(MCG/)[0-9]*\\s*(ML)" # str Pattern 1: x MCG/ML
conc_pattern_2 <- "[0-9]+\\s(MG/)[0-9]*\\s*(ML)" # str Pattern 2: x MG/ML
conc_pattern_3 <- "[0-9]+\\s(MG/)[0-9]+\\.*[0-9]*\\s*(ML)" # str Pattern 3: x MG/y ML 
conc_pattern_4 <- "[0-9]+\\s(MCG/)[0-9]+\\.*[0-9]*\\s*(ML)" # str Pattern 4: x MCG/y ML 
conc_pattern_5 <- "[0-9]+\\s(MG\\s)" # str Pattern 5: x MG
conc_pattern_6 <- "[0-9]*,*[0-9]+\\s(UNIT/)[0-9]*\\s*(ML)" # str Pattern 6: x UNITS/y ML

covid_meds_admin <- covid_meds_admin %>%
  mutate(StrExtract = ifelse(!is.na(str_extract(DISPINSABLE_MED_NAME, conc_pattern_1)), 
                             str_extract(DISPINSABLE_MED_NAME, conc_pattern_1),
                             ifelse(!is.na(str_extract(DISPINSABLE_MED_NAME, conc_pattern_2)),
                                    str_extract(DISPINSABLE_MED_NAME, conc_pattern_2),
                                    ifelse(!is.na(str_extract(DISPINSABLE_MED_NAME, conc_pattern_3)),
                                           str_extract(DISPINSABLE_MED_NAME, conc_pattern_3),
                                           ifelse(!is.na(str_extract(DISPINSABLE_MED_NAME, conc_pattern_4)),
                                                  str_extract(DISPINSABLE_MED_NAME, conc_pattern_4),
                                                  ifelse(!is.na(str_extract(DISPINSABLE_MED_NAME, conc_pattern_5)),
                                                         str_extract(DISPINSABLE_MED_NAME, conc_pattern_5), 
                                                         ifelse(!is.na(str_extract(DISPINSABLE_MED_NAME, conc_pattern_6)),
                                                                str_extract(DISPINSABLE_MED_NAME, conc_pattern_6), NA)))))))

# Replace /ML with /1 ML
covid_meds_admin <- covid_meds_admin %>%
  mutate(StrExtract = str_replace(StrExtract, "/ML", "/1 ML"))



# split StrExtract
covid_meds_admin$StrExtract_split <- covid_meds_admin$StrExtract
covid_meds_admin <- covid_meds_admin %>% separate(StrExtract_split, c("a", "b"), extra = "merge", fill = "right", sep = "/")
covid_meds_admin <- covid_meds_admin %>%  separate(a, c( "ConcDoseSize", "ConcDoseUnit"), extra = "merge", fill = "right", sep = " ")
covid_meds_admin <- covid_meds_admin %>%  separate(b, c( "ConcVolSize", "ConcVolUnit"), extra = "merge", fill = "right", sep = " ")


covid_meds_admin <- covid_meds_admin %>% mutate(ConcDoseUnit= str_trim(ConcDoseUnit, side = c("both")))


covid_meds_admin <- covid_meds_admin %>% mutate(ConcDoseSize=gsub(",","", ConcDoseSize),
                                                ConcVolSize= gsub(",","", ConcVolSize) )

covid_meds_admin <- covid_meds_admin %>% mutate(ConcDoseSize=  as.numeric(as.character(ConcDoseSize)),
                                                ConcVolSize=  as.numeric(as.character(ConcVolSize)))



# Capitalize all units
covid_meds_admin$MAR_DOSE_UNITS <- toupper(covid_meds_admin$MAR_DOSE_UNITS)
covid_meds_admin$ADMIN_UNIT <- toupper(covid_meds_admin$ADMIN_UNIT)
covid_meds_admin$ORDER_VOLUME_UNIT <- toupper(covid_meds_admin$ORDER_VOLUME_UNIT)
covid_meds_admin$ConcDoseUnit <- toupper(covid_meds_admin$ConcDoseUnit)
covid_meds_admin$ConcVolUnit <- toupper(covid_meds_admin$ConcVolUnit)

# Format to standardize units 
covid_meds_admin$ConcDoseUnit[which(covid_meds_admin$ConcDoseUnit == "UNIT")] <- "UNITS"

# Normalized doses 
covid_meds_admin <- covid_meds_admin %>%  mutate(NormDoseSize = ifelse(is.na(ConcVolSize), ConcDoseSize, ConcDoseSize / ConcVolSize),
                                                 NormDoseUnit = paste0(ConcDoseUnit,"/",ConcVolUnit)) %>% mutate(NormDoseUnit = ifelse(NormDoseUnit == "/", NA,
                                                                                                                                       ifelse(NormDoseUnit == "MG/", "MG", NormDoseUnit)))
covid_meds_admin$NormDoseSize <- as.numeric(covid_meds_admin$NormDoseSize)

# Normalized concentration per ml 
covid_meds_admin <- covid_meds_admin %>%  mutate(NormConcPerMl = ifelse(ConcVolUnit == "ML", ConcDoseSize / ConcVolSize, ""))
covid_meds_admin$NormConcPerMl <- as.numeric(covid_meds_admin$NormConcPerMl)


# Calculate total doses administered
covid_meds_admin <- covid_meds_admin %>%
  mutate(total_doses = ifelse(MAR_DOSE_UNITS %in% c("MG","MCG","UNITS"), MAR_DOSE,
                              ifelse(ADMIN_UNIT %in% c("MG","MCG","UNITS"), ADMIN_AMOUNT, 
                                     ifelse(ADMIN_UNIT == "ML" & !is.na(ConcDoseUnit), ADMIN_AMOUNT*NormConcPerMl,""))))

covid_meds_admin$total_doses[is.na(covid_meds_admin$total_doses)] <- ""

covid_meds_admin <- covid_meds_admin %>%
  mutate(total_doses = ifelse(total_doses == "" & ORDER_VOLUME_UNIT == "ML" & ConcVolUnit == "ML" & !is.na(ConcDoseUnit),
                              ORDER_VOLUME*NormConcPerMl, total_doses))
covid_meds_admin$total_doses[is.na(covid_meds_admin$total_doses)] <- ""
covid_meds_admin$total_doses <- as.numeric(covid_meds_admin$total_doses)

# Get correct units for total doses administered 
covid_meds_admin <- covid_meds_admin %>%
  mutate(total_doses_unit = ifelse(MAR_DOSE_UNITS %in% c("MG","MCG","UNITS"), MAR_DOSE_UNITS,
                                   ifelse(ADMIN_UNIT %in% c("MG","MCG","UNITS"), ADMIN_UNIT, 
                                          ifelse(ADMIN_UNIT == "ML" & !is.na(ConcDoseUnit), ConcDoseUnit,""))))

covid_meds_admin$total_doses_unit[is.na(covid_meds_admin$total_doses_unit)] <- ""

covid_meds_admin <- covid_meds_admin %>%
  mutate(total_doses_unit = ifelse(total_doses_unit == "" & ORDER_VOLUME_UNIT == "ML" & ConcVolUnit == "ML" & !is.na(ConcDoseUnit),
                                   ConcDoseUnit, total_doses_unit))
covid_meds_admin$total_doses_unit[is.na(covid_meds_admin$total_doses_unit)] <- ""


# Site rolllup
covid_meds_admin <- covid_meds_admin %>% mutate(loc_rollup= ifelse(str_detect(LOC_NAME, "BETH ISRAEL| BI PETRIE_DEACTIVATED"), "MSBI",
                                                                   ifelse(str_detect(LOC_NAME, "MORNINGSIDE|ST LUKE'S_DEACTIVATED"),"MSM",
                                                                          ifelse(str_detect(LOC_NAME,"HOSPITAL|MADISON AVE"), "MSH",
                                                                                 ifelse(str_detect(LOC_NAME, "BROOKLYN"), "MSB",
                                                                                        ifelse(str_detect(LOC_NAME, "QUEENS"), "MSQ",
                                                                                               ifelse(str_detect(LOC_NAME,"WEST"), "MSW", NA )))))))



covid_meds_admin <- covid_meds_admin %>% filter(!is.na(loc_rollup ))


# Bind today's data with repository
med_final_repo <- rbind(med_repo, covid_meds_admin)
med_final_repo <- med_final_repo %>% distinct()



# Save the new repo
#write_csv(med_final_repo, path = paste0(wrk.dir, "\\REPO\\Emergency_Med_Repo\\Covid Surge Meds Admin Repo-", Sys.Date(), ".csv"))

saveRDS(med_final_repo  ,paste0("J:/deans/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Pharmacy/Data/Epic Latest Reports/Daily Reports/REPO/Emergency_Med_Repo/Covid Surge Meds Admin Repo-", Sys.Date(), ".RDS"))


rm(med_repo, new_med_admin, covid_meds_admin, emergency_med_groups)



setwd(wrk.dir)
setwd("../../..")

#setwd("C:\\Users\\aghaer01\\Downloads\\Code")

save_output <- paste0(getwd(), "\\Daily Reporting Output")
rmarkdown::render("Code\\Emergency-Meds-Report-Rmarkdown-2022-01-18.Rmd", 
                  output_file = paste("Emergency Meds Report-", Sys.Date()), output_dir = save_output)
