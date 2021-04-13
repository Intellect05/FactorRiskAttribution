
## Time series of daily, weekly or monthly asset returns with the first column being the date 
Stocks<-read_excel("\\path...file.xlsx", trim_ws = FALSE, col_names = TRUE, na ="")[,-1] ##drop Date Column

## Time series of daily, weekly or monthly factor returns with the first column being the date 
Factors<-read_excel(""\\path...file.xlsx"", trim_ws = FALSE, col_names = TRUE, na ="")[,-1] ##drop Date Column

## Weights of the assets in the portfolio and index
Weights<-read_excel(""\\path...file.xlsx"", trim_ws = FALSE, col_names = TRUE, na ="")

source("r scrip path\\script.R")

##Multi-Factor-Risk Model illustration
MultiFactor_Attr(Stocks = Stocks, Factors = Factors,PortfolioWeights = Weights[,4],IndexWeights = Weights[,5],ActiveWeights = Weights[,6],
                 method =c("Factor risk attribution"), scale = 52, drop_factor =0)
