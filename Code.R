library(xlsx)
library(readxl)
library(hydroGOF)
library(randomForest)
library(iml)
library(ggplot2)
library(circlize)
library(dplyr)
library(randomForestExplainer)
library(iml)
library(tcltk)
library(patchwork)
library(raster)
library(ggbreak)
library(grid)
library(ggpointdensity)
library(openxlsx)
library(writexl)
setwd('E:/Phd/NP-plant')
label <- c(excel_sheets('data.xlsx'))
#------------------------------------------------------------------------------#
#-----------------------------------RF model-----------------------------------#
#------------------------------------------------------------------------------#
Test_predict_wb <- createWorkbook()
Test_predict_10fold_wb <- createWorkbook()
Train_predict_10fold_wb <- createWorkbook()
for (i in 1:length(label)){
  addWorksheet(Test_predict_wb,sheetName = label[i])
}
for (i in 1:length(label)){
  addWorksheet(Test_predict_10fold_wb,sheetName = label[i])
}
for (i in 1:length(label)){
  addWorksheet(Train_predict_10fold_wb,sheetName = label[i])
}
for (i in 1:length(label)){
  data_model <- read.xlsx('data.xlsx', i)
  set.seed(1234);disorder <- sample(nrow(data_model),replace=F)
  fold_num <- floor(nrow(data_model)/10)
  
  n_test <- data.frame()
  n_train <- data.frame()
  predict <- data.frame()
  
  for (k in 1:10){
    o <- disorder[(fold_num*(k-1)+1):(fold_num*k)]
    rf.data <- data_model[-o,]
    rf <- randomForest(index~. , data = rf.data,
                       ntree=500 ,mtry=20,
                       proximity = F,
                       importance = F)
    p <- as.data.frame(predict(rf, data_model[o,]))
    a <- as.data.frame(data_model[o,ncol(data_model)])
    p_train <- as.data.frame(predict(rf, rf.data))
    a_train <- as.data.frame(rf.data[,ncol(data_model)])
    
    r2 <- cor(p,a)
    R2 <- R2(p,a)
    rm <- Metrics::rmse(p[,1],a[,1])
    predict <- rbind(predict,cbind(a,p))
    n_test <- rbind(n_test,cbind(r2, R2, rm))
    #train
    r2_train <- cor(p_train,a_train)
    R2_train <- R2(p_train,a_train)
    rm_train <- Metrics::rmse(p_train[,1],a_train[,1])
    n_train <- rbind(n_train,cbind(r2_train, R2_train, rm_train))
    
    print(k)
  }
  colnames(predict) <- c('Observation','Prediction')
  colnames(n_test) <- c('r2','R2','RMSE')
  colnames(n_train) <- c('r2','R2','RMSE')
  
  writeData(Test_predict_wb, sheet = i, predict)
  writeData(Test_predict_10fold_wb, sheet = i, n_test)
  writeData(Train_predict_10fold_wb, sheet = i, n_train)
  print(i)
  print(cor(predict))
}
saveWorkbook(Test_predict_wb, "Test_predict.xlsx", overwrite = TRUE)
saveWorkbook(Test_predict_10fold_wb, "Test_predict_10fold.xlsx", overwrite = TRUE)
saveWorkbook(Train_predict_10fold_wb, "Train_predict_10fold.xlsx", overwrite = TRUE)


