
options(warn=-1)



suppressPackageStartupMessages({
  library(tseries)
  library(xlsx)
  library(readxl)
  library("tibble")
  library("janitor")
  require(pls)
  require(ridge)
})


MultiFactor_Attr<-function(Stocks, Factors,Port_returns, BM_returns, PortfolioWeights=rep(1,ncol(Stocks)),IndexWeights=rep(1,ncol(Stocks)),
                           ActiveWeights=rep(1,ncol(Stocks)),Portfolio_name ,Model = c("MLR Model", "PCR Model", "PLSR Model","LR Model"),
                           method=c("x sigma rho attribution","Alpha Beta attribution","MCTR","Factor risk attribution",
                                    "Factor MCTR", "Exposure Analysis","VaR","Risk Ratios","Performance ratios"), 
                           scale=NA,drop_factor_1=NA,drop_factor_2=NA,drop_factor_3=NA,drop_factor_4=NA,drop_factor_5=NA,
                           drop_factor_6=NA,drop_factor_7=NA,drop_factor_8=NA,drop_factor_9=NA,drop_factor_10=NA,drop_factor_11=NA,
                           drop_factor_12=NA,drop_factor_13=NA,drop_factor_14=NA,drop_factor_15=NA,drop_factor_16=NA,drop_factor_17=NA,
                           drop_factor_18=NA,drop_factor_19=NA,drop_factor_20=NA,drop_factor_21=NA,tail.prob,h=NA,Date, output_Path){
  
  if(is.na(drop_factor_1)){
    drop_factor_1<-as.numeric(drop_factor_1)}
  
  if(is.na(drop_factor_2)){
    drop_factor_2<-as.numeric(drop_factor_2)}
  
  if(is.na(drop_factor_3)){
    drop_factor_3<-as.numeric(drop_factor_3)}
  
  if(is.na(drop_factor_4)){
    drop_factor_4<-as.numeric(drop_factor_4)}
  
  if(is.na(drop_factor_5)){
    drop_factor_5<-as.numeric(drop_factor_5)}
  
  if(is.na(drop_factor_6)){
    drop_factor_6<-as.numeric(drop_factor_6)}
  
  if(is.na(drop_factor_7)){
    drop_factor_7<-as.numeric(drop_factor_7)}
  
  if(is.na(drop_factor_8)){
    drop_factor_8<-as.numeric(drop_factor_8)}
  
  if(is.na(drop_factor_9)){
    drop_factor_9<-as.numeric(drop_factor_9)}
  
  if(is.na(drop_factor_10)){
    drop_factor_10<-as.numeric(drop_factor_10)}
  
  if(is.na(drop_factor_11)){
    drop_factor_11<-as.numeric(drop_factor_11)}
  
  if(is.na(drop_factor_12)){
    drop_factor_12<-as.numeric(drop_factor_12)}
  
  if(is.na(drop_factor_13)){
    drop_factor_13<-as.numeric(drop_factor_13)}
  
  if(is.na(drop_factor_14)){
    drop_factor_14<-as.numeric(drop_factor_14)}
  
  if(is.na(drop_factor_15)){
    drop_factor_15<-as.numeric(drop_factor_15)}
  
  if(is.na(drop_factor_16)){
    drop_factor_16<-as.numeric(drop_factor_16)}
  
  if(is.na(drop_factor_17)){
    drop_factor_17<-as.numeric(drop_factor_17)}
  
  if(is.na(drop_factor_18)){
    drop_factor_18<-as.numeric(drop_factor_18)}
  
  if(is.na(drop_factor_19)){
    drop_factor_19<-as.numeric(drop_factor_19)}
  
  if(is.na(drop_factor_20)){
    drop_factor_20<-as.numeric(drop_factor_20)}
  
  if(is.na(drop_factor_21)){
    drop_factor_21<-as.numeric(drop_factor_21)}
  
  
  if(is.na(h)){
    h<-as.integer(h)}
  
  Port_returns<-ts(Port_returns)
  BM_returns<-ts(BM_returns)
  
  if(nrow(Port_returns)!= nrow(BM_returns)){
    message("Portfolio returns should have the same number of rows with index returns")
  }
  
  list(Stocks, all.names = TRUE)
  
  Stocks<-as.data.frame(Stocks)
  
  
  for( i in 1:nrow(Stocks)){
    for( j in 1:ncol(Stocks)){
      Stocks[,j]=Stocks[,j]-mean(Stocks[,j])
    }
  }
  
  
  list(Factors, all.names = TRUE)
  
  Factors<-as.data.frame(Factors)
  
  
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
  
  ## Option to Drop certian factors 
  
  if(drop_factor_1>0){
    Factors[,drop_factor_1]<-0
  }
  
  if(drop_factor_2>0){
    Factors[,drop_factor_2]<-0
  }
  
  if(drop_factor_3>0){
    Factors[,drop_factor_3]<-0
  }
  
  if(drop_factor_4>0){
    Factors[,drop_factor_4]<-0
  }
  
  if(drop_factor_5>0){
    Factors[,drop_factor_5]<-0
  }
  
  if(drop_factor_6>0){
    Factors[,drop_factor_6]<-0
  }
  
  if(drop_factor_7>0){
    Factors[,drop_factor_7]<-0
  }
  
  if(drop_factor_8>0){
    Factors[,drop_factor_8]<-0
  }
  
  if(drop_factor_9>0){
    Factors[,drop_factor_9]<-0
  }
  
  if(drop_factor_10>0){
    Factors[,drop_factor_10]<-0
  }
  
  if(drop_factor_11>0){
    
    Factors[,drop_factor_11]<-0
  }
  
  if(drop_factor_12>0){
    Factors[,drop_factor_12]<-0
  }
  
  if(drop_factor_13>0){
    Factors[,drop_factor_13]<-0
  }
  
  if(drop_factor_14>0){
    Factors[,drop_factor_14]<-0
  }
  
  if(drop_factor_15>0){
    Factors[,drop_factor_15]<-0
  }
  
  if(drop_factor_16>0){
    Factors[,drop_factor_16]<-0
  }
  
  if(drop_factor_17>0){
    Factors[,drop_factor_17]<-0
  }
  
  if(drop_factor_18>0){
    Factors[,drop_factor_18]<-0
  }
  
  if(drop_factor_19>0){
    Factors[,drop_factor_19]<-0
  }
  
  if(drop_factor_20>0){
    Factors[,drop_factor_20]<-0
  }
  
  if(drop_factor_21>0){
    Factors[,drop_factor_21]<-0
  }

  ##Calc Factor Covariance Matrix
  
  F_Cov<-cov(Factors, use= "everything")
  
  
  if(var(Factors[,1]) != F_Cov[1,1]){
    message("the variances in the factor returns does not match the variances in the Factor Covariance Matrix ")
  }
  
  

  if( Model == "MLR Model"){
    output=matrix(nrow=ncol(Factors)+1, ncol=ncol(Stocks))
  for(i in 1: length(Stocks)){
    
    for( j in 1:i){   
      R <- lm(Stocks[,i]~., data = as.data.frame(Factors))   
    }
    
    output[,i:ncol(Stocks)]<-rep(R$coefficients[1:(ncol(Factors)+1)])
    
    colnames(output) <- colnames(Stocks)
    rownames(output)<- c("intercept", colnames(Factors))
    
   }
  }
  
  
  if( Model == "PCR Model"){
    
    output=matrix(nrow=ncol(Factors), ncol=ncol(Stocks))
    
    for(i in 1: length(Stocks)){
      
      for( j in 1:i){   
        pcr_model <- pcr(Stocks[,i]~., data = as.data.frame(Factors), scale =FALSE, validation = "CV")  
      }
      
      output[,i:ncol(Stocks)]<-rep(pcr_model$coefficients[1:ncol(Factors)])
      
      colnames(output) <- colnames(Stocks)
      rownames(output)<- c(colnames(Factors))
      
    }
  }
  
  if(Model == "PLSR Model"){
    
    output=matrix(nrow=ncol(Factors), ncol=ncol(Stocks))
    
    for(i in 1: length(Stocks)){
      
      for( j in 1:i){   
        plsr_model <- plsr(Stocks[,i]~., data = as.data.frame(Factors), ncomps=10, validation = "LOO")  
      }
      
      output[,i:ncol(Stocks)]<-rep(plsr_model$coefficients[1:ncol(Factors)])
      
      colnames(output) <- colnames(Stocks)
      rownames(output)<- c(colnames(Factors))
      
    }
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
    
    return(results = write.xlsx(xSigmaRho,paste0("",as.character(output_Path),"x sigma rho attribution.xlsx")))
  }
  
  ##Calculate 
  
  if(Model == "PCR Model" || Model == "PLSR Model"){
  Beta_returns<-as.matrix(Factors)%*%as.matrix(output[1:nrow(output),1:ncol(output)])
  } else{Beta_returns<-as.matrix(Factors)%*%as.matrix(output[2:nrow(output),1:ncol(output)]) }
  
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
    
    return(results = write.xlsx(Alpha_Beta_attr,paste0("",as.character(output_Path),"Alpha beta attribution.xlsx")))
    
  }
  
  if( method == "VaR"){
    
    return(results = write.xlsx(VaR,paste0("",as.character(output_Path),"VaR.xlsx")))
  }
  
  if( method =="MCTR"){
    
    
    Systematic_MCTR<-((as.matrix(cov(Beta_returns, use = "everything"))%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(cov(Beta_returns, use = "everything"))%*%as.matrix(ActiveWeights)))))*sqrt(scale)
    
    Systematic_CCTR<-ActiveWeights*Systematic_MCTR
    
    Specific_MCTR<-((as.matrix(cov(Alpha_returns, use = "everything"))%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(cov(Alpha_returns, use = "everything"))%*%as.matrix(ActiveWeights)))))*sqrt(scale)
    
    Specific_CCTR<-ActiveWeights*Specific_MCTR
    
    MCTR=data.frame(Date,ActiveWeights, Systematic_MCTR,Specific_MCTR,Systematic_CCTR,Specific_CCTR)
    
    colnames(MCTR)<-c("Date","Active Weights","Systematic MCTR","Specific MCTR","Systematic CCTR","Specific CCTR")
    
    
    return(results = write.xlsx(MCTR,paste0("",as.character(output_Path),"MCTR.xlsx.xlsx")))
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
    
    return(results = write.xlsx(Exposure,paste0("",as.character(output_Path),"Exposure Analysis.xlsx")))
  }
  
  F_Matrices=matrix(nrow=ncol(Factors), ncol=ncol(Factors))
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,1])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),1))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),1))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),1))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor1<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor1<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  
  
  for( i in 1:length(Factors)){
    for ( j in 1:i){
      
      F_Matrices[i,]<-(F_Cov[i,2])
      
      if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),2))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),2))]<0){
        F_Matrices[i,as.vector(setdiff(1:ncol(Factors),2))]<-0
      }
      if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor2<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))
      }
      else{Factor2<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
      }
    }
  }
  
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,3])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),3))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),3))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),3))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor3<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))
      }
    else{Factor3<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
    }
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,4])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),4))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),4))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),4))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor4<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))
      }
    else{Factor4<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))
    }
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,5])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),5))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),5))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),5))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor5<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor5<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,6])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),6))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),6))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),6))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor6<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor6<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,7])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),7))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),7))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),7))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor7<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor7<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,8])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),8))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),8))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),8))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor8<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor8<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,9])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),9))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),9))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),9))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor9<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor9<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,10])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),10))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),10))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),10))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor10<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor10<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    
    F_Matrices[i,]<-(F_Cov[i,11])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),11))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),11))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),11))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor11<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor11<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,12])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),12))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),12))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),12))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor12<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor12<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,13])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),13))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),13))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),13))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor13<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor13<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,14])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),14))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),14))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),14))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor14<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor14<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  

  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,15])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),15))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),15))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),15))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor15<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor15<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,16])
  
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),16))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),16))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),16))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor16<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor16<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,17])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),17))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),17))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),17))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor17<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor17<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,18])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),18))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),18))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),18))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor18<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor18<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,19])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),19))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),19))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),19))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor19<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor19<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }

  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,20])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),20))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),20))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),20))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor20<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor20<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,21])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),21))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),21))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),21))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor21<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor21<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,22])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),22))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),22))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),22))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor22<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor22<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,23])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),23))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),23))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),23))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor23<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor23<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,24])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),24))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),24))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),24))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor24<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor24<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,25])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),25))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),25))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),25))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor25<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor25<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,26])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),26))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),26))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),26))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor26<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor26<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,27])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),27))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),27))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),27))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor27<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor27<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,28])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),28))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),28))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),28))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor28<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor28<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,29])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),29))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),29))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),29))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor29<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor29<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,30])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),30))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),30))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),30))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor30<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor30<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,31])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),31))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),31))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),31))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor31<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    
    else{Factor31<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,32])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),32))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),32))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),32))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor32<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor32<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,33])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),33))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),33))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),33))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor33<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor33<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,34])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),34))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),34))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),34))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor34<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor34<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,35])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),35))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),35))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),35))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor35<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor35<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,36])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),36))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),36))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),36))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor36<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor36<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  for( i in 1:length(Factors)){
    F_Matrices[i,]<-(F_Cov[i,37])
    
    if(F_Matrices[i,as.vector(setdiff(1:ncol(Factors),37))]>0|F_Matrices[i,as.vector(setdiff(1:ncol(Factors),37))]<0){
      F_Matrices[i,as.vector(setdiff(1:ncol(Factors),37))]<-0
    }
    if(Model == "PCR Model" || Model == "PLSR Model"){
      Factor37<-t(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[1:nrow(output),1:ncol(Stocks)]))}
    else{Factor37<-t(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))%*%as.matrix(F_Matrices)%*%(as.matrix(output[2:nrow(output),1:ncol(Stocks)]))}
  }
  
  
  Total_Factor =(as.matrix(Factor1)+as.matrix(Factor2)+as.matrix(Factor3)+as.matrix(Factor4)+as.matrix(Factor5)+as.matrix(Factor6)+as.matrix(Factor7)+as.matrix(Factor8)+as.matrix(Factor9)+as.matrix(Factor10)+as.matrix(Factor11)+as.matrix(Factor12)
                 +as.matrix(Factor13)+as.matrix(Factor14)+as.matrix(Factor15)+as.matrix(Factor16)+as.matrix(Factor17)+as.matrix(Factor18)+as.matrix(Factor19)+as.matrix(Factor20)+as.matrix(Factor21)+as.matrix(Factor22)+as.matrix(Factor23)
                 +as.matrix(Factor24)+as.matrix(Factor25)+as.matrix(Factor26)+as.matrix(Factor27)+as.matrix(Factor28)+as.matrix(Factor29)+as.matrix(Factor30)+as.matrix(Factor31)+as.matrix(Factor32)+as.matrix(Factor33)+as.matrix(Factor34)
                 +as.matrix(Factor35)+as.matrix(Factor36)+as.matrix(Factor37))
  
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
  
  F14_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor14)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F15_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor15)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F16_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor16)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F17_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor17)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F18_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor18)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F19_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor19)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F20_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor20)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F21_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor21)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F22_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor22)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F23_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor23)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F24_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor24)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F25_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor25)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F26_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor26)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F27_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor27)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F28_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor28)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F29_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor29)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F30_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor30)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F31_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor31)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F32_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor32)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F33_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor33)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F34_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor34)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F35_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor35)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F36_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor36)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  F37_risk=((t(as.matrix(ActiveWeights))%*%as.matrix(Factor37)%*%as.matrix(ActiveWeights))/Total_risk1)*sqrt(scale)
  
  
  Total_Factor_risk=((F1_risk+F2_risk+F3_risk+F4_risk+F5_risk+F6_risk+F7_risk+F8_risk+F9_risk+F10_risk+F11_risk+F12_risk+F13_risk+F14_risk+F15_risk+F16_risk+F17_risk+F18_risk+F19_risk+F20_risk
                      +F21_risk+F22_risk+F23_risk+F24_risk+F25_risk+F26_risk+F27_risk+F28_risk+F29_risk+F30_risk+F31_risk+F32_risk+F33_risk+F34_risk+F35_risk+F36_risk+F37_risk))*sqrt(scale)
  
  Total_r <-data.frame(Date,F1_risk,F2_risk,F3_risk,F4_risk,F5_risk,F6_risk,F7_risk,F8_risk,F9_risk,F10_risk,F11_risk,F12_risk,F13_risk,F14_risk,F15_risk,F16_risk,F17_risk,F18_risk,
                       F19_risk,F20_risk,F21_risk,F22_risk,F23_risk,F24_risk,F25_risk,F26_risk,F27_risk,F28_risk,F29_risk,F30_risk,F31_risk,F32_risk,F33_risk,F34_risk,
                       F35_risk,F36_risk,F37_risk,Total_Factor_risk)
  
  colnames(Total_r)<-c("Date",colnames(Factors),"Total Factor Risk")
  row.names(Total_r)<-c("Factor risk attribution")
  
  if( method == "Factor risk attribution"){
    
    return(results = write.xlsx(Total_r,paste0("",as.character(output_Path),"Factor risk attribution.xlsx")))
  }
  
  ##Factor Att  

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
  
  ##Factor14
  Factor14_MCTR<-((as.matrix(Factor14)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor14_CCTR<-ActiveWeights*Factor14_MCTR
  
  ##Factor15
  Factor15_MCTR<-((as.matrix(Factor15)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor15_CCTR<-ActiveWeights*Factor15_MCTR
  
  ##Factor16
  Factor16_MCTR<-((as.matrix(Factor16)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor16_CCTR<-ActiveWeights*Factor16_MCTR
  
  ##Factor17
  Factor17_MCTR<-((as.matrix(Factor17)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor17_CCTR<-ActiveWeights*Factor17_MCTR
  
  ##Factor18
  Factor18_MCTR<-((as.matrix(Factor18)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor18_CCTR<-ActiveWeights*Factor18_MCTR
  
  ##Factor19
  Factor19_MCTR<-((as.matrix(Factor19)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor19_CCTR<-ActiveWeights*Factor19_MCTR
  
  ##Factor20
  Factor20_MCTR<-((as.matrix(Factor20)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor20_CCTR<-ActiveWeights*Factor20_MCTR
  
  ##Factor21
  Factor21_MCTR<-((as.matrix(Factor21)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor21_CCTR<-ActiveWeights*Factor21_MCTR
  
  ##Factor22
  Factor22_MCTR<-((as.matrix(Factor22)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor22_CCTR<-ActiveWeights*Factor22_MCTR
  
  ##Factor23
  Factor23_MCTR<-((as.matrix(Factor23)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor23_CCTR<-ActiveWeights*Factor23_MCTR
  
  ##Factor24
  Factor24_MCTR<-((as.matrix(Factor24)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor24_CCTR<-ActiveWeights*Factor24_MCTR
  
  ##Factor25
  Factor25_MCTR<-((as.matrix(Factor25)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor25_CCTR<-ActiveWeights*Factor25_MCTR
  
  ##Factor26
  Factor26_MCTR<-((as.matrix(Factor26)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor26_CCTR<-ActiveWeights*Factor26_MCTR
  
  ##Factor27
  Factor27_MCTR<-((as.matrix(Factor27)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor27_CCTR<-ActiveWeights*Factor27_MCTR
  
  ##Factor28
  Factor28_MCTR<-((as.matrix(Factor28)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor28_CCTR<-ActiveWeights*Factor28_MCTR
  
  ##Factor29
  Factor29_MCTR<-((as.matrix(Factor29)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor29_CCTR<-ActiveWeights*Factor29_MCTR
  
  ##Factor30
  Factor30_MCTR<-((as.matrix(Factor30)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor30_CCTR<-ActiveWeights*Factor30_MCTR
  
  ##Factor31
  Factor31_MCTR<-((as.matrix(Factor31)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor31_CCTR<-ActiveWeights*Factor31_MCTR
  
  ##Factor32
  Factor32_MCTR<-((as.matrix(Factor32)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor32_CCTR<-ActiveWeights*Factor32_MCTR
  
  ##Factor33
  Factor33_MCTR<-((as.matrix(Factor33)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor33_CCTR<-ActiveWeights*Factor33_MCTR
  
  ##Factor34
  Factor34_MCTR<-((as.matrix(Factor34)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor34_CCTR<-ActiveWeights*Factor34_MCTR
  
  ##Factor22
  Factor35_MCTR<-((as.matrix(Factor35)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor35_CCTR<-ActiveWeights*Factor35_MCTR
  
  ##Factor36
  Factor36_MCTR<-((as.matrix(Factor36)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor36_CCTR<-ActiveWeights*Factor36_MCTR
  
  ##Factor37
  Factor37_MCTR<-((as.matrix(Factor37)%*%as.matrix(ActiveWeights))%*%(1/sqrt((t(as.matrix(ActiveWeights))%*%as.matrix(Total_Factor)%*%as.matrix(ActiveWeights)))))*sqrt(scale)
  
  Factor37_CCTR<-ActiveWeights*Factor37_MCTR
  
  
  Factor_MCTR_CCTR=data.frame(Date,ActiveWeights, Factor1_CCTR,Factor2_CCTR,Factor3_CCTR,Factor4_CCTR,Factor5_CCTR,Factor6_CCTR,Factor7_CCTR,Factor8_CCTR,Factor9_CCTR,
                              Factor10_CCTR,Factor11_CCTR,Factor12_CCTR,Factor13_CCTR,Factor14_CCTR,Factor15_CCTR,Factor16_CCTR,Factor17_CCTR,Factor18_CCTR,Factor19_CCTR,
                              Factor20_CCTR,Factor21_CCTR,Factor22_CCTR,Factor23_CCTR,Factor24_CCTR,Factor25_CCTR,Factor26_CCTR,Factor27_CCTR,Factor28_CCTR,Factor29_CCTR,
                              Factor30_CCTR,Factor31_CCTR,Factor32_CCTR,Factor33_CCTR,Factor34_CCTR,Factor35_CCTR,Factor36_CCTR,Factor37_CCTR)
  
  
  colnames(Factor_MCTR_CCTR)<-c("Date","Acitve Weights",colnames(Factors))
  rownames(Factor_MCTR_CCTR)<-colnames(Stocks)
  
  if(method == "Factor MCTR"){
    
    return(results = write.xlsx(Factor_MCTR_CCTR,paste0("",as.character(output_Path),"Factor MCTR.xlsx")))
    
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
    
    return(results = write.xlsx(table.PR,paste0("",as.character(output_Path),"Performance_ratios.xlsx")))
  }
  
  table.RR<-data.frame(Date,Portfolio_volatility,BM_volatility,TE,IR,SR,SortinoR, TreynorR, Upside_risk,Downside_risk,
                       Ulcer_Index,Calmar_Ratio)
  
  colnames(table.RR)<-c("Date","Portfolio Volatility ", "BM Volatility", "Ex Post TE","IR","SR","Sortino Ratio","Treynor Ratio",
                        "Upside Risk","Downside Risk","Ulcer Index","Calmar Ratio")
  
  rownames(table.RR)<-Portfolio_name
  
  if(method == "Risk Ratios"){
    
    return(results = write.xlsx(table.RR,paste0("",as.character(output_Path),"Risk Ratios.xlsx")))
  }
  
  return(results)
}
