
options(warn=-1)

suppressPackageStartupMessages({
  library(tseries)
  library(xlsx)
  library(readxl)
  library("tibble")
  library("janitor")
})


MultiFactor_Attr<-function(Stocks, Factors,Port_returns, BM_returns, PortfolioWeights=rep(1,ncol(Stocks)),IndexWeights=rep(1,ncol(Stocks)),
                           ActiveWeights=rep(1,ncol(Stocks)),Portfolio_name ,
                           method=c("x sigma rho attribution","Alpha Beta attribution","MCTR","Factor risk attribution",
                                    "Factor MCTR", "Exposure Analysis","VaR","Risk Ratios","Performance ratios"), 
                           scale=NA,drop_factor_1=NA,drop_factor_2=NA,drop_factor_3=NA,tail.prob,h=NA,Date){
  
  if(is.na(drop_factor_1)){
    drop_factor_1<-as.numeric(drop_factor_1)}
  
  if(is.na(drop_factor_2)){
    drop_factor_2<-as.numeric(drop_factor_2)}
  
  if(is.na(drop_factor_3)){
    drop_factor_3<-as.numeric(drop_factor_3)}

  if(is.na(h)){
    h<-as.integer(h)}
  
  Port_returns<-ts(Port_returns)
  BM_returns<-ts(BM_returns)
  
  if(nrow(Port_returns)!= nrow(BM_returns)){
    message("Portfolio returns should have the same number of rows with index returns")
  }
  
  Stocks<-as.data.frame(Stocks)
  
  list(Stocks, all.names = TRUE)
  
  for( i in 1:nrow(Stocks)){
    for( j in 1:ncol(Stocks)){
      Stocks[,j]=Stocks[,j]-mean(Stocks[,j])
    }
  }
  
  Factors<-as.data.frame(Factors)
  
  list(Factors, all.names = TRUE)
  
  Factors[] <- lapply(Factors, function(x) as.numeric(as.character(x)))
  
  for( i in 1:nrow(Factors)){
    for( j in 1:ncol(Factors)){
      Factors[,j]=Factors[,j]-mean(Factors[,j])
    }
  }
  
  if(nrow(Stocks)!=nrow(Factors)){
    message("The timeseries length of asset returns should be equal to that factor returns")
  }
  
  # check Active weights options
  if (!is.null(ActiveWeights)) {
    if (is.vector(ActiveWeights)){
      # message("weights are a vector, will use same weights for entire time series") # remove this warning if you call function recursively
      if (nrow(ActiveWeights)!=ncol(Stocks)) {
        stop("number of items in weighting vector not equal to number of columns in Stocks")
      }
    } else {
      ActiveWeights<-as.data.frame(ActiveWeights)
      list(ActiveWeights, all.names = TRUE)
      if (nrow(ActiveWeights) != ncol(Stocks)) {
        stop("number of columns in weighting timeseries not equal to number of columns in Stocks")
      }
      #@todo: check for date overlap with R and weights
    }
  } # end weight checks
  
  # check weights options
  if (!is.null(IndexWeights)) {
    if (is.vector(IndexWeights)){
      # message("weights are a vector, will use same weights for entire time series") # remove this warning if you call function recursively
      if (nrow(IndexWeights)!=ncol(Stocks)) {
        stop("number of items in weighting vector not equal to number of columns in Stocks")
      }
    } else {
      IndexWeights<-as.data.frame(IndexWeights)
      list(IndexWeights, all.names = TRUE)
      if (nrow(IndexWeights) != ncol(Stocks)) {
        stop("number of columns in weighting timeseries not equal to number of columns in Stocks")
      }
      #@todo: check for date overlap with R and weights
    }
  } # end IndexWeights checks
  
  # check weights options
  if (!is.null(PortfolioWeights)) {
    if (is.vector(PortfolioWeights)){
      # message("weights are a vector, will use same weights for entire time series") # remove this warning if you call function recursively
      if (nrow(PortfolioWeights)!=ncol(Stocks)) {
        stop("number of items in weighting vector not equal to number of columns in Stocks")
      }
    } else {
      PortfolioWeights<-as.data.frame(PortfolioWeights)
      list(PortfolioWeights, all.names = TRUE)
      if (nrow(PortfolioWeights) != ncol(Stocks)) {
        stop("number of columns in weighting timeseries not equal to number of columns in Stocks")
      }
      #@todo: check for date overlap with R and weights
    }
  } # end IndexWeights checks
  
  #Specify the scale of the returns, i.e., daily, weekly, monthly, quarterly
  if(is.na(scale) && !xtsible(x))
    stop("'x' needs to be timeBased or xtsible, or scale must be specified." )
  
  if(is.na(scale)) {
    freq = periodicity(x)
    switch(freq$scale,
           minute = {stop("Data periodicity too high")},
           hourly = {stop("Data periodicity too high")},
           daily = {scale = 252},
           weekly = {scale = 52},
           monthly = {scale = 12},
           quarterly = {scale = 4},
           yearly = {scale = 1}
    )
  }
  
  if(drop_factor_1>0){
    Factors[,drop_factor_1]<-0
  }
  
  if(drop_factor_2>0){
    Factors[,drop_factor_2]<-0
  }
  
  if(drop_factor_3>0){
    Factors[,drop_factor_3]<-0
  }
  
  ##Calc Factor Covariance Matrix
  
  F_Cov<-cov(Factors, use= "everything")
  
  if(var(Factors[,1]) != F_Cov[1,1]){
    message("the variances in the factor returns does not match the variances in the Factor Covariance Matrix ")
  }
  
  output=matrix(nrow=ncol(Factors)+1, ncol=ncol(Stocks))
  for(i in 1: length(Stocks)){
    
    for( j in 1:i){   
      R <- lm(Stocks[,i]~., data = as.data.frame(Factors))   
    }
    
    output[,i:ncol(Stocks)]<-rep(R$coefficients[1:(ncol(Factors)+1)])
    
    colnames(output) <- colnames(Stocks)
    rownames(output)<- c("intercept", colnames(Factors))
    
  }
  
  for( i in 1:nrow(output)){
    for( j in 1:ncol(output)){
      if(is.na(output)[i,j]){output[i,j]<-0}
    }
  }
  
  if(method == "x sigma rho attribution"){
    x_sigma_rho= (sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(cov(Stocks, use = "everything"))%*%as.matrix(ActiveWeights)))*sqrt(scale))
    
    xSigmaRho = data.frame(Date,x_sigma_rho)
    colnames(xSigmaRho)<-c("Date","Ex ante Tracking Error")
    rownames(xSigmaRho)<-c(Portfolio_name)
    

    return(results=write.xlsx(xSigmaRho,"U:\\Risk_Model\\x sigma rho attribution.xlsx"))
  }
  
  ##Calculate 
  Beta_returns<-as.matrix(Factors)%*%as.matrix(output[2:nrow(output),1:ncol(output)])
  
  Alpha_returns = Stocks[]-Beta_returns
  
  Specific_risk = (sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(cov(Alpha_returns, use = "everything"))%*%as.matrix(ActiveWeights)))*sqrt(scale))
  
  
  Systematic_risk=(sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(cov(Beta_returns, use = "everything"))%*%as.matrix(ActiveWeights)))*sqrt(scale))
  
  Total_risk=sqrt(sum(Systematic_risk^2+Specific_risk^2))
  
  if ( tail.prob < 0 || tail.prob > 1){
    stop("tail.prob must be between 0 and 1")}
  
  Systematic_VaR =(-Systematic_risk* qnorm(tail.prob)*sqrt(as.integer(h)/scale))
  Specific_VaR =(-Specific_risk* qnorm(tail.prob)*sqrt(as.integer(h)/scale))
  Total_VaR = (-Total_risk*qnorm(tail.prob)*sqrt(as.integer(h)/scale))
  
  Alpha_Beta_attr<-data.frame(Date,Specific_risk,Systematic_risk,Total_risk)
  
  colnames(Alpha_Beta_attr)<-c("Date","Specific Risk", "Systematic Risk", "Total Risk")
  
  rownames(Alpha_Beta_attr)<-Portfolio_name
  
  VaR<- data.frame(Date,Systematic_VaR,Specific_VaR,Total_VaR)
  colnames(VaR)<-c("Date","Systematic VaR","Specific VaR", "Total VaR")
  rownames(VaR)<-Portfolio_name
  
  if( method == "Alpha Beta attribution"){
    
    return(results=write.xlsx(Alpha_Beta_attr,"U:\\Risk_Model\\Alpha beta attribution.xlsx"))}
  
  if( method == "VaR"){
    return(results =write.xlsx(VaR,"U:\\Risk_Model\\VaR.xlsx"))
  }
  
  if( method =="MCTR"){
    
    Beta_returns<-as.matrix(Factors)%*%as.matrix(output[2:nrow(output),1:ncol(output)])
    
    Alpha_returns = Stocks[]-Beta_returns
    
    Systematic_MCTR<-((as.matrix(cov(Beta_returns, use = "everything"))%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(cov(Beta_returns, use = "everything"))%*%as.matrix(ActiveWeights)))))*sqrt(scale)
    
    Systematic_CCTR<-ActiveWeights*Systematic_MCTR
    
    Specific_MCTR<-((as.matrix(cov(Alpha_returns, use = "everything"))%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(cov(Alpha_returns, use = "everything"))%*%as.matrix(ActiveWeights)))))*sqrt(scale)
    
    Specific_CCTR<-ActiveWeights*Specific_MCTR
    
    MCTR=data.frame(Date,ActiveWeights, Systematic_MCTR,Specific_MCTR,Systematic_CCTR,Specific_CCTR)
    
    colnames(MCTR)<-c("Date","Active Weights","Systematic MCTR","Specific MCTR","Systematic CCTR","Specific CCTR")
    
    return(results=write.xlsx(MCTR,"U:\\Risk_Model\\MCTR.xlsx"))
  }
  
  if(method == "Exposure Analysis"){
    Port_exposure = matrix(nrow=nrow(output), ncol = 1)
    BM_exposure = matrix(nrow=nrow(output), ncol = 1)
    
    for(i in 1:nrow(output)){
      
      Port_exposure[i,]<-t(as.matrix(output[i,1:ncol(output)]))%*%as.matrix(PortfolioWeights)
      BM_exposure[i,]<-t(as.matrix(output[i,1:ncol(output)]))%*%as.matrix(IndexWeights)
      Active_exposure=(Port_exposure-BM_exposure)
      Fund_Exposure<-data.frame(Date,sum(Port_exposure[2:nrow(Port_exposure),]),sum(BM_exposure[2:nrow(BM_exposure),]),
                                sum(Active_exposure[2:nrow(Port_exposure),]))
      Factor_Exposure<-data.frame(Date,Port_exposure,BM_exposure,Active_exposure)
      
      colnames(Factor_Exposure)<-c("Date","Portfolio", "BM", "Active")
      rownames(Factor_Exposure)<-rownames(output)
      
      colnames(Fund_Exposure)<-c("Date","Portfolio", "BM", "Active")
      rownames(Fund_Exposure)<-Portfolio_name
      
      Exposure<-rbind.data.frame(Fund_Exposure,Factor_Exposure)
      
    }
    return(results = write.xlsx(Exposure,"U:\\Risk_Model\\Exposure Analysis.xlsx"))
  }
  
  F_Matrices=matrix(nrow=ncol(Factors), ncol=ncol(Factors))
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,1])
    
    if(F_Matrices[i,2:ncol(F_Matrices)]>0|F_Matrices[i,2:ncol(F_Matrices)]<0){
      F_Matrices[i,2:ncol(F_Matrices)]<-0
    }
    Factor1<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    for ( j in 1:i){
      
      F_Matrices[i,]<-(F_Cov[i,2])
      
      if(F_Matrices[i,c(1,3:ncol(F_Matrices))]>0|F_Matrices[i,c(1,3:ncol(F_Matrices))]<0){
        F_Matrices[i,c(1,3:ncol(F_Matrices))]<-0
      }
      Factor2<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
    }
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,3])
    
    if(F_Matrices[i,c(1,2,4:ncol(F_Matrices))]>0|F_Matrices[i,c(1,2,4:ncol(F_Matrices))]<0){
      F_Matrices[i,c(1,2,4:ncol(F_Matrices))]<-0
    }
    Factor3<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,4])
    
    if(F_Matrices[i,c(1,2,3,5:ncol(F_Matrices))]>0|F_Matrices[i,c(1,2,3,5:ncol(F_Matrices))]<0){
      F_Matrices[i,c(1,2,3,5:ncol(F_Matrices))]<-0
    }
    Factor4<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,5])
    
    if(F_Matrices[i,c(1,2,3,4,6:ncol(F_Matrices))]>0|F_Matrices[i,c(1,2,3,4,6:ncol(F_Matrices))]<0){
      F_Matrices[i,c(1,2,3,4,6:ncol(F_Matrices))]<-0
    }
    Factor5<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,6])
    
    if(F_Matrices[i,c(1,2,3,4,5,7:ncol(F_Matrices))]>0|F_Matrices[i,c(1,2,3,4,5,7:ncol(F_Matrices))]<0){
      F_Matrices[i,c(1,2,3,4,5,7:ncol(F_Matrices))]<-0
    }
    Factor6<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,7])
    
    if(F_Matrices[i,c(1,2,3,4,5,6,8:ncol(F_Matrices))]>0|F_Matrices[i,c(1,2,3,4,5,6,8:ncol(F_Matrices))]<0){
      F_Matrices[i,c(1,2,3,4,5,6,8:ncol(F_Matrices))]<-0
    }
    Factor7<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,8])
    
    if(F_Matrices[i,c(1,2,3,4,5,6,7,9:ncol(F_Matrices))]>0|F_Matrices[i,c(1,2,3,4,5,6,7,9:ncol(F_Matrices))]<0){
      F_Matrices[i,c(1,2,3,4,5,6,7,9:ncol(F_Matrices))]<-0
    }
    Factor8<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,9])
    
    if(F_Matrices[i,c(1,2,3,4,5,6,7,8,10:ncol(F_Matrices))]>0|F_Matrices[i,c(1,2,3,4,5,6,7,8,10:ncol(F_Matrices))]<0){
      F_Matrices[i,c(1,2,3,4,5,6,7,8,10:ncol(F_Matrices))]<-0
    }
    Factor9<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,10])
    
    if(F_Matrices[i,c(1,2,3,4,5,6,7,8,9,11:ncol(F_Matrices))]>0|F_Matrices[i,c(1,2,3,4,5,6,7,8,9,11:ncol(F_Matrices))]<0){
      F_Matrices[i,c(1,2,3,4,5,6,7,8,9,11:ncol(F_Matrices))]<-0
    }
    Factor10<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,11])
    
    if(F_Matrices[i,c(1,2,3,4,5,6,7,8,9,10,12,13)]>0|F_Matrices[i,c(1,2,3,4,5,6,7,8,9,10,12,13)]<0){
      F_Matrices[i,c(1,2,3,4,5,6,7,8,9,10,12,13)]<-0
    }
    Factor11<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,12])
    
    if(F_Matrices[i,c(1,2,3,4,5,6,7,8,9,10,11,13:ncol(F_Matrices))]>0|F_Matrices[i,c(1,2,3,4,5,6,7,8,9,11,13:ncol(F_Matrices))]<0){
      F_Matrices[i,c(1,2,3,4,5,6,7,8,9,10,11,13:ncol(F_Matrices))]<-0
    }
    Factor12<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,13])
    
    if(F_Matrices[i,c(1,2,3,4,5,6,7,8,9,10,11,12)]>0|F_Matrices[i,c(1,2,3,4,5,6,7,8,9,11,12)]<0){
      F_Matrices[i,c(1,2,3,4,5,6,7,8,9,10,11,12)]<-0
    }
    Factor13<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
  }
  
  Total_Factor =(as.matrix(Factor1)+as.matrix(Factor2)+as.matrix(Factor3)+as.matrix(Factor4)+as.matrix(Factor5)+as.matrix(Factor6)+as.matrix(Factor7)+as.matrix(Factor8)+as.matrix(Factor9)+as.matrix(Factor10)+as.matrix(Factor11)+as.matrix(Factor12)
                 +as.matrix(Factor13))
  
  Total_risk1=sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))*sqrt(scale)
  
  F1_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor1)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F2_risk = ((t(as.matrix(ActiveWeights))%*%as.matrix(Factor2)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F3_risk = ((t(as.matrix(ActiveWeights))%*%as.matrix(Factor3)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F4_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor4)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F5_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor5)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F6_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor6)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F7_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor7)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F8_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor8)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F9_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor9)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F10_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor10)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F11_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor11)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F12_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor12)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F13_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor13)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  Total_Factor_risk=((F1_risk+F2_risk+F3_risk+F4_risk+F5_risk+F6_risk+F7_risk+F8_risk+F9_risk+F10_risk+F11_risk+F12_risk+F13_risk)*sqrt(scale))
  
  Total_r <-data.frame(Date,F1_risk,F2_risk,F3_risk,F4_risk,F5_risk,F6_risk,F7_risk,F8_risk,F9_risk,F10_risk,F11_risk,F12_risk,F13_risk,Total_Factor_risk)
  colnames(Total_r)<-c("Date",colnames(Factors),"Total Factor Risk")
  row.names(Total_r)<-c("Factor risk attribution")
  
  if( method == "Factor risk attribution"){
    return(results =  write.xlsx(Total_r,"U:\\Risk_Model\\Factor risk attribution.xlsx"))
  }
  
  ##Factor Attribution
  
  ##Factor1
  Factor1_MCTR<-((as.matrix(Factor1)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor1_CCTR<-ActiveWeights*Factor1_MCTR
  
  ##Factor2
  Factor2_MCTR<-((as.matrix(Factor2)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor2_CCTR<-ActiveWeights*Factor2_MCTR
  
  ##Factor3
  
  Factor3_MCTR<-((as.matrix(Factor3)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor3_CCTR<-ActiveWeights*Factor3_MCTR
  
  ##Factor4
  Factor4_MCTR<-((as.matrix(Factor4)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor4_CCTR<-ActiveWeights*Factor4_MCTR
  
  ##Factor5
  Factor5_MCTR<-((as.matrix(Factor5)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor5_CCTR<-ActiveWeights*Factor5_MCTR
  
  ##Factor6
  Factor6_MCTR<-((as.matrix(Factor6)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor6_CCTR<-ActiveWeights*Factor6_MCTR
  
  ##Factor7
  Factor7_MCTR<-((as.matrix(Factor7)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor7_CCTR<-ActiveWeights*Factor7_MCTR
  
  ##Factor8
  Factor8_MCTR<-((as.matrix(Factor8)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor8_CCTR<-ActiveWeights*Factor8_MCTR
  
  ##Factor9
  Factor9_MCTR<-((as.matrix(Factor9)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor9_CCTR<-ActiveWeights*Factor9_MCTR
  
  ##Factor10
  Factor10_MCTR<-((as.matrix(Factor10)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor10_CCTR<-ActiveWeights*Factor10_MCTR
  
  ##Factor11
  Factor11_MCTR<-((as.matrix(Factor11)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor11_CCTR<-ActiveWeights*Factor11_MCTR
  
  ##Factor12
  Factor12_MCTR<-((as.matrix(Factor12)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor12_CCTR<-ActiveWeights*Factor12_MCTR
  
  ##Factor13
  Factor13_MCTR<-((as.matrix(Factor13)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor13_CCTR<-ActiveWeights*Factor13_MCTR
  
  Factor_MCTR_CCTR=data.frame(Date,ActiveWeights, Factor1_CCTR,Factor2_CCTR,Factor3_CCTR,Factor4_CCTR,Factor5_CCTR,Factor6_CCTR,Factor7_CCTR,Factor8_CCTR,Factor9_CCTR,
                              Factor10_CCTR,Factor11_CCTR,Factor12_CCTR,Factor13_CCTR)
  colnames(Factor_MCTR_CCTR)<-c("Date","Acitve Weights",colnames(Factors))
  rownames(Factor_MCTR_CCTR)<-colnames(Stocks)
  
  if(method == "Factor MCTR"){
    return(results=write.xlsx(Factor_MCTR_CCTR,"U:\\Risk_Model\\Factor MCTR.xlsx"))
  }
  
  Portfolio_return<-(as.matrix(Stocks)%*%as.matrix(PortfolioWeights))
  
  Portfolio_return.annualized =  prod(1 + Port_returns) - 1
  
  BM_return<-(as.matrix(Stocks)%*%as.matrix(IndexWeights))
  
  BM_return.annualized =  prod(1 + BM_returns) - 1
  
  Active_returns<- Portfolio_return.annualized-BM_return.annualized
  
  Portfolio_volatility=StdDev.annualized(Port_returns, scale = scale)
  
  BM_volatility=StdDev.annualized(BM_returns, scale = scale)
  
  ##TE =(sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(cov(Stocks, use = "everything"))%*%as.matrix(ActiveWeights)))*sqrt(scale))
  
  TE = TrackingError(Port_returns,BM_returns, scale = scale)
  
  ##IR = Active_returns/TE
  
  IR = InformationRatio(Port_returns,BM_returns, scale = scale)
  
  SR=SharpeRatio(Port_returns, FUN = "StdDev")
  
  SortinoR =SortinoRatio(Port_returns)
  
  TreynorR = TreynorRatio(Port_returns,BM_returns, scale = scale)
  
  Upside_risk = UpsideRisk(Port_returns, method = c("full"), stat = c("risk"))
  
  Downside_risk =DownsideDeviation(Port_returns, method = c("full"))
  
  Ulcer_Index= UlcerIndex(Port_returns)
  
  Calmar_Ratio <-CalmarRatio(Port_returns, scale = scale)
  
  UpDown_Ratios = UpDownRatios(Port_returns,BM_returns,method = c("Capture"))
  
  Port_MaxDrawdown <- -maxDrawdown(Port_returns)
  
  BM_MaxDrawdown<- -maxDrawdown(BM_returns)
  
  table.PR<-data.frame(Date,Portfolio_return.annualized,BM_return.annualized,Active_returns,
                       Port_MaxDrawdown,BM_MaxDrawdown,UpDown_Ratios[1],UpDown_Ratios[2])
  
  colnames(table.PR)<-c("Date","Portfolio return (ann)", "BM return (ann)", "Active return", 
                        "Portfolio max-drawdon", "BM max-drawdown","Up Capture Ratio","Down Capture Ratio")
  
  rownames(table.PR)<-Portfolio_name
  
  if(method == "Performance ratios"){
    
    return(results = write.xlsx(table.PR,"U:\\Risk_Model\\Performance_ratios.xlsx"))
  }
  
  table.RR<-data.frame(Date,Portfolio_volatility,BM_volatility,TE,IR,SR,SortinoR, TreynorR, Upside_risk,Downside_risk,
                       Ulcer_Index,Calmar_Ratio)
  
  colnames(table.RR)<-c("Date","Portfolio Volatility ", "BM Volatility", "Ex Post TE","IR","SR","Sortino Ratio","Treynor Ratio",
                        "Upside Risk","Downside Risk","Ulcer Index","Calmar Ratio")
  
  rownames(table.RR)<-Portfolio_name
  
  if(method == "Risk Ratios"){
    
    return(results = write.xlsx(table.RR,"U:\\Risk_Model\\Risk Ratios.xlsx"))
  }
  
  return(results)
}
