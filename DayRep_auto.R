library(readxl)
library(dplyr)
library(openxlsx)

args <- commandArgs(trailingOnly = FALSE)
script_path <- NULL

if (any(grepl("--file=", args))) {
  script_path <- normalizePath(sub("--file=", "", args[grep("--file=", args)]))
}

if (is.null(script_path) || script_path == "") {
  script_dir <- getwd()
} else {
  script_dir <- dirname(script_path)
}

data_directory <- file.path(script_dir)
prefix <- "Statement"
files <- list.files(data_directory, pattern = paste0("^", prefix, ".*\\.xlsx$"), full.names = TRUE)




#######################################################################

data <- read_excel(files)

N <- 13
data <- data[-(1:N), ]

comp_names_col <- data[[1]]
type_col <- data[["Type"]]

colnames(data) <- data[1, ]
data <- data[-(1:1), ]

company_names <- unique(data[[1]])

purchase <- unique(comp_names_col[type_col == "purchase"])
payout <- unique(comp_names_col[type_col == "payout"])

currency <- unique(data[["Currency"]])
pay_meth <- unique(data[["Payment Method"]])

total_purch <- data.frame(Name = character(), Currency = character(), Count = numeric(), Paid = numeric(), stringsAsFactors = FALSE)
total_payo <- data.frame(Name = character(), Currency = character(), Count = numeric(), Success = numeric(), stringsAsFactors = FALSE)


data[["Amount"]] <- as.numeric(data[["Amount"]])

count_meth_vector <- c()
turn_purch_vector <- c()
turn_payo_vector <- c()
turn_purch_vector_DE <- c()

for (name in company_names) {
  for (curr in currency) {
      count_total <- sum(data[[1]] == name & data[["Type"]] == "purchase" & data[["Currency"]] == curr)
      if (count_total != 0) {
        count_paid <- sum(data[[1]] == name & data[["Type"]] == "purchase" & data[["Currency"]] == curr & data[["Status"]] == "paid")
        
        purch_data_filt <- data %>%
           filter(`Legal Name` == name, Type == "purchase", Currency == curr, Status == "paid")
        turn_purch <- sum(purch_data_filt[["Amount"]], na.rm = TRUE)
        turn_purch_vector <- c(turn_purch_vector, turn_purch)
        
        purch_data_filt_DE <- data %>%
          filter(`Legal Name` == name, Type == "purchase", Currency == curr, Status == "paid", Country == "DE")
        turn_purch_DE <- sum(purch_data_filt_DE[["Amount"]], na.rm = TRUE)
        turn_purch_vector_DE <- c(turn_purch_vector_DE, turn_purch_DE)
        
        count_meth <- sum(data[[1]] == name & data[["Type"]] == "purchase" & data[["Currency"]] == curr & data[["Status"]] == "paid" & data[["Payment Method"]] %in% c("mastercard", "maestro"))
        count_meth_vector <- c(count_meth_vector, count_meth)
        total_purch <- rbind(total_purch, data.frame(Name = name, Currency = curr, Count = count_total, Paid = count_paid))
      }
    
      count_total <- sum(data[[1]] == name & data[["Type"]] == "payout" & data[["Currency"]] == curr)
      if (count_total != 0) {
        
        payo_data_filt <- data %>%
          filter(`Legal Name` == name, Type == "payout", Currency == curr, Status == "success")
        turn_payo <- sum(payo_data_filt[["Amount"]], na.rm = TRUE)
        turn_payo_vector <- c(turn_payo_vector, turn_payo)
        
        count_suc <- sum(data[[1]] == name & data[["Type"]] == "payout" & data[["Currency"]] == curr & data[["Status"]] == "success")
        total_payo <- rbind(total_payo, data.frame(Name = name, Currency = curr, Count = count_total, Success = count_suc))  
      }
  }
}

total_purch$err <- total_purch[["Count"]] - total_purch[["Paid"]]
total_payo$err <- total_payo[["Count"]] - total_payo[["Success"]]

total_purch$paid_p <- (total_purch[["Paid"]]/total_purch[["Count"]]) * 100
total_payo$suc_p <- (total_payo[["Success"]]/total_payo[["Count"]]) * 100
total_purch$paid_p <- sprintf("%.2f%%", total_purch$paid_p)
total_payo$suc_p <- sprintf("%.2f%%", total_payo$suc_p)

total_purch$err_p <- (total_purch[["err"]]/total_purch[["Count"]]) * 100
total_payo$err_p <- (total_payo[["err"]]/total_payo[["Count"]]) * 100
total_purch$err_p <- sprintf("%.2f%%", total_purch$err_p)
total_payo$err_p <- sprintf("%.2f%%", total_payo$err_p)

total_purch$Mastercard <- count_meth_vector

total_purch$Turnover <- turn_purch_vector
total_purch$Turnover_DE <- turn_purch_vector_DE
total_payo$Turnover <- turn_payo_vector


total_purch_card <- total_purch[total_purch$Mastercard != 0, ]
total_purch_noncard <- total_purch[total_purch$Mastercard == 0, ]

rownames(total_purch_card) <- seq_len(nrow(total_purch_card))
rownames(total_purch_noncard) <- seq_len(nrow(total_purch_noncard))
total_purch_noncard$Mastercard <- NULL

total_purch_card$Visa <- total_purch_card$Paid - total_purch_card$Mastercard
total_purch_card$Mastercard_perc <- (total_purch_card$Mastercard/total_purch_card$Paid) * 100
total_purch_card$Visa_perc <- (total_purch_card$Visa/total_purch_card$Paid) * 100

total_purch_card$Mastercard_perc <- sprintf("%.2f%%", total_purch_card$Mastercard_perc)
total_purch_card$Visa_perc <- sprintf("%.2f%%", total_purch_card$Visa_perc)

total_purch_card <- total_purch_card %>%
  select(-Turnover, Turnover) %>%  
  select(-Turnover_DE, Turnover_DE)


######################################################################

file_name <- file.path(script_dir, "output.xlsx")
wb <- createWorkbook()

addWorksheet(wb, "total_purch_card")
addWorksheet(wb, "total_purch_noncard")
addWorksheet(wb, "total_payo")

writeData(wb, sheet = 1, total_purch_card)
writeData(wb, sheet = 2, total_purch_noncard)
writeData(wb, sheet = 3, total_payo)

saveWorkbook(wb, file = file_name, overwrite = TRUE)