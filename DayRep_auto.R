library(readxl)
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

total_purch <- data.frame(Name = character(), Currency = character(), Count = numeric(), stringsAsFactors = FALSE)
total_payo <- data.frame(Name = character(), Currency = character(), Count = numeric(), stringsAsFactors = FALSE)

for (name in company_names) {
  for (curr in currency) {
    count <- sum(data[[1]] == name & data[["Type"]] == "purchase" & data[["Currency"]] == curr)
    total_purch <- rbind(total_purch, data.frame(Name = name, Currency = curr, Count = count))
    
    count <- sum(data[[1]] == name & data[["Type"]] == "payout" & data[["Currency"]] == curr)
    total_payo <- rbind(total_payo, data.frame(Name = name, Currency = curr, Count = count))
  }
}
total_purch <- total_purch[total_purch$Count != 0, ]
total_payo <- total_payo[total_payo$Count != 0, ]

