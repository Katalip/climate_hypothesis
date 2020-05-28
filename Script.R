# Setting up Working directory
setwd('C:/Users/katal/Documents/Kudaibergen/Study/6 semester/R Stat/Final')
getwd()

# Library for working with excel files
library("openxlsx")

file_names <- c('AQI', 'CO', 'NO2', 'O3', 'PM2.5', 'PM10', 'SO2')
iters <- 1
for (file_name in file_names) { # Reading every file using for loop
  n <- sprintf('Data/%s.csv', file_name)
  df <- read.csv(n) # Reading
  
  for_saving_t <- data.frame(City = character(), Statistic = numeric(0), 
                             Parameter = numeric(0), p.value = numeric(0),
                             conf.int = character(),  difference_in_means = numeric(0),
                             alternative = character(),
                             stringsAsFactors=FALSE)
  
  for_saving_f <- data.frame(City = character(), Statistic = numeric(0), 
                             Parameter = numeric(0), p.value = numeric(0),
                             conf.int = character(),  ratio_of_variances = numeric(0),
                             alternative = character(),
                             stringsAsFactors=FALSE)
  
  
  # 169:181 indices of cities located in Hubei province
  # 9:31 are indices for 01.01.20 - 23.01.20 (mean values per day)
  # 69:91 are indices for 01.03.20 - 23.03.20 (mean values per day)
  # 399:421, 459:481 are same just for variances per day
  
  for(i in c(169:181)) {
    a <- t.test(t(df[i, 9:31]), t(df[i, 69:91]), paired = TRUE) #, alternative = 'greater') # t - transpose
    b <- var.test(t(df[i, 399:421]^2), t(df[i, 459:481]^2))
    
    
    for_saving_t[nrow(for_saving_t)+1, ] <- c(sprintf('%s',df[i, 'City_EN']), a$statistic, a$parameter, a$p.value, 
                                              sprintf('%s, %s',a$conf.int[1], a$conf.int[2]), a$estimate, 
                                              a$alternative)
    
    for_saving_f[nrow(for_saving_f)+1, ] <- c(sprintf('%s',df[i, 'City_EN']), b$statistic, b$parameter[1], b$p.value, 
                                              sprintf('%s, %s',b$conf.int[1], b$conf.int[2]), b$estimate, 
                                              b$alternative)
    
    
  }
  
  a <- t.test(t(df[169:181, 9:31]), t(df[169:181, 69:91]), paired = TRUE) # For full province
  b <- var.test(t(df[169:181, 399:421]^2), t(df[169:181, 459:481]^2)) # For full province
  
  
  a <- data.frame(Statistic = a$statistic, 
                  Parameter = a$parameter, p.value = a$p.value,
                  conf.int = sprintf('%s, %s',a$conf.int[1], a$conf.int[2]),  difference_in_means = a$estimate,
                  alternative = a$alternative,
                  stringsAsFactors=FALSE)
  
  b <- data.frame(Statistic = b$statistic, 
                  Parameter = b$parameter[1], p.value = b$p.value,
                  conf.int = sprintf('%s, %s',b$conf.int[1], b$conf.int[2]),  ratio_of_variances = b$estimate,
                  alternative = b$alternative,
                  stringsAsFactors=FALSE)
  
  
  if(iters == 1) wb <- createWorkbook('new_test')
  else wb <- loadWorkbook("new_test.xlsx")
  
  addWorksheet(wb, sprintf("t_%s", file_name))
  addWorksheet(wb, sprintf("t_full_%s", file_name))
  
  addWorksheet(wb, sprintf("f_%s", file_name))
  addWorksheet(wb, sprintf("f_full_%s", file_name))
  
  writeData(wb, sprintf("t_%s", file_name), for_saving_t, startRow = 1, startCol = 1)
  writeData(wb, sprintf("f_%s", file_name), for_saving_f, startRow = 1, startCol = 1)
  writeData(wb, sprintf("t_full_%s", file_name), a, startRow = 1, startCol = 1)
  writeData(wb, sprintf("f_full_%s", file_name), b, startRow = 1, startCol = 1)
  
  saveWorkbook(wb, file = "new_test.xlsx", overwrite = TRUE)
  iters <- iters + 1
  
}
