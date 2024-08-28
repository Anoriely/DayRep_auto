library(readxl)
library(dplyr)

data <- read_excel("/Users/dmitrijsciruks/DayReport/Statement (2024-08-28).xlsx")

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

for (name in company_names) {
  for (curr in currency) {
      count_total <- sum(data[[1]] == name & data[["Type"]] == "purchase" & data[["Currency"]] == curr)
      count_paid <- sum(data[[1]] == name & data[["Type"]] == "purchase" & data[["Currency"]] == curr & data[["Status"]] == "paid")
      total_purch <- rbind(total_purch, data.frame(Name = name, Currency = curr, Count = count_total, Paid = count_paid))
      
      count_total <- sum(data[[1]] == name & data[["Type"]] == "payout" & data[["Currency"]] == curr)
      count_suc <- sum(data[[1]] == name & data[["Type"]] == "payout" & data[["Currency"]] == curr & data[["Status"]] == "success")
      total_payo <- rbind(total_payo, data.frame(Name = name, Currency = curr, Count = count_total, Success = count_suc)) 
  }
}

total_purch <- total_purch[total_purch$Count != 0, ]
total_payo <- total_payo[total_payo$Count != 0, ]

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

