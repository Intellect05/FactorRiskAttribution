
options(warn=-1)

suppressPackageStartupMessages({
  library(tseries)
  library(xlsx)
  library(readxl)
  library(PerformanceAnalytics)
  require(pls)
  require(ridge)
})


## This is the settings workbook from the dashboard remember to specify the location of file in the VBA code 

Risk_Data<-read.csv("U:\\Risk_Model\\Risk_Data.csv")

Stocks<-read_excel(paste0(as.character(Risk_Data[,"Stocks"][1])), trim_ws = TRUE, col_names = TRUE, na ="")[,-1]


Returns<-read_excel(as.character(Risk_Data[,"Returns"][1]), trim_ws = TRUE, col_names = TRUE, na ="")


Factors<-read_excel(as.character(Risk_Data[,"Factors"][1]), trim_ws = FALSE, col_names = TRUE, na ="")[,-1]


Weights<-read_excel(as.character(Risk_Data[,"Weights"][1]), trim_ws = FALSE, col_names = TRUE, na ="")


source("U:\\Risk_Model\\Risk_Model_RScript.R")

##Remember to index the right columns for the portfolio and index in the weights and returns data templates i.e., if Portfolio are in column 6 then in the function Port_returns = Returns_data[,6] etc.

MultiFactor_Attr(Stocks = Stocks, Factors = Factors,Port_returns = Returns[,9],BM_returns = Returns[,11],PortfolioWeights = Weights[,4],IndexWeights = Weights[,5],ActiveWeights = Weights[,6],
                 Portfolio_name= as.character(Risk_Data[,"Portfolio"][1]),Model =c(as.character(Risk_Data[,"Model"][1])), method =c(as.character(Risk_Data[,"Method"][1])), scale = as.numeric(Risk_Data[,"Data_frequency"][1]),
                 drop_factor_1  =(Risk_Data[,"drop_factor_1"][1]),drop_factor_2 = (Risk_Data[,"drop_factor_2"][1]),drop_factor_3 =(Risk_Data[,"drop_factor_3"][1]),
                 drop_factor_4 =(Risk_Data[,"drop_factor_4"][1]),drop_factor_5 =(Risk_Data[,"drop_factor_5"][1]),drop_factor_6 =(Risk_Data[,"drop_factor_6"][1]),drop_factor_7 =(Risk_Data[,"drop_factor_7"][1]),
                 drop_factor_8 =(Risk_Data[,"drop_factor_8"][1]),drop_factor_9 =(Risk_Data[,"drop_factor_9"][1]),drop_factor_10 =(Risk_Data[,"drop_factor_10"][1]),
                 drop_factor_11 =(Risk_Data[,"drop_factor_11"][1]),drop_factor_12 =(Risk_Data[,"drop_factor_12"][1]),drop_factor_13 =(Risk_Data[,"drop_factor_13"][1]),
                 drop_factor_14 =(Risk_Data[,"drop_factor_14"][1]),drop_factor_15 =(Risk_Data[,"drop_factor_15"][1]),drop_factor_16 =(Risk_Data[,"drop_factor_16"][1]),
                 drop_factor_17 =(Risk_Data[,"drop_factor_17"][1]),drop_factor_18 =(Risk_Data[,"drop_factor_18"][1]),drop_factor_19 =(Risk_Data[,"drop_factor_19"][1]),
                 drop_factor_20 =(Risk_Data[,"drop_factor_20"][1]),drop_factor_21 =(Risk_Data[,"drop_factor_21"][1]),
                 tail.prob = as.vector(Risk_Data[,"tail_probability"][1]), h=as.integer(Risk_Data[,"Number_of_Days"][1]), 
                 Date = excel_numeric_to_date(as.numeric(as.character(Risk_Data[,"Date"][1])), date_system = "modern"),output_Path = Risk_Data[,"output"][1])
