

Stocks<-read_excel("\\path...file.xlsx", trim_ws = FALSE, col_names = TRUE, na ="")[,-1]

Factors<-read_excel(""\\path...file.xlsx"", trim_ws = FALSE, col_names = TRUE, na ="")[,-1]

Weights<-read_excel(""\\path...file.xlsx"", trim_ws = FALSE, col_names = TRUE, na ="")


source("r scrip path\\script.R")

MultiFactor_Attr(Stocks = Stocks, Factors = Factors,PortfolioWeights = Weights[,4],IndexWeights = Weights[,5],ActiveWeights = Weights[,6],
                 method =c("Factor risk attribution"), scale = 52, drop_factor =0)
