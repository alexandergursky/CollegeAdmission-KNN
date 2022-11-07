
# College Admission using KNN method


library(readxl)
library(caret)

myData <- read_xlsx("College Admissions-2Case5.xlsx")
myDataScore <- read_xlsx("College Admissions-2Case5.xlsx", sheet = 2)
#View(myData)
#View(myDataScore)


myData1 <- subset(myData, myData$College == "Business & Economics")
#View(myData1)


myData1$Dummy_GenderF <- ifelse(myData1$Gender == 'F', 1, 0)
myData1$Dummy_AdmittedYes <- ifelse(myData1$Admitted =='Yes', 1, 0)

myData1 <- myData1[ -c(1, 4, 9, 10, 11, 12)]


myData.st <- data.frame(scale(myData1[c(1:6)]), myData1[c(7:8)])

#Changing Dummies Factors
str(myData.st)
myData.st$Dummy_AdmittedYes <- as.factor(myData.st$Dummy_AdmittedYes)
myData.st$Dummy_GenderF <- as.factor(myData.st$Dummy_GenderF)

str(myData.st)


set.seed(1)
myIndex <- createDataPartition(myData.st$Dummy_AdmittedYes, p=0.7, list = FALSE)
trainSet <- myData.st[myIndex,]
validationSet <- myData.st[-myIndex,]

#Creating a table
accuracyTable <- data.frame(k=seq(3,12,1), Accuracy = NA, Sensitivity = NA,
                            Specificity = NA)

for (i in 3:12)
{
  myGrid <- data.frame(k=i)
  set.seed(1)
  KNN_Fit <- train(Dummy_AdmittedYes ~ ., data = trainSet, method = 'knn', tuneGrid = myGrid)
  KNN_Fit
  KNN_Class <- predict(KNN_Fit, newdata = validationSet)
  #Imputing accuracy
  accuracyTable[i-2, 2] <- confusionMatrix(KNN_Class, validationSet$Dummy_AdmittedYes,
                                           positive = '1')$overall[1]
  #Imputing sensitivity
  accuracyTable[i-2, 3] <- confusionMatrix(KNN_Class, validationSet$Dummy_AdmittedYes,
                                           positive = '1')$byClass[1]
  #Imputing specificity
  accuracyTable[i-2, 4] <- confusionMatrix(KNN_Class, validationSet$Dummy_AdmittedYes,
                                           positive = '1')$byClass[2]
  
}
#output <- confusionMatrix(KNN_Class, validationSet$Dummy_AdmittedYes,positive = '1')
#output
#output$overall
#output$byClass
#output$byClass[1] # Sensitivity
#output$byClass[2] # Specificity



myDataScore$Dummy_GenderF <- ifelse(myDataScore$Gender == 'F', 1, 0)
myDataScore$Dummy_GenderF <- as.factor(myDataScore$Dummy_GenderF)
myDataScore <- myDataScore[-c(4)]
myDataScore <- myDataScore[-c(1)]
myDataScore.st <- data.frame(scale(myDataScore[c(1:6)]), myDataScore[c(7)])


# Rerunning the optimal KNN model


myGrid1 <- data.frame(k=11)
set.seed(1)
KNN_Fit1 <- train(Dummy_AdmittedYes ~ ., data = trainSet, method = 'knn', tuneGrid = myGrid1)
KNN_Fit1
KNN_Class1 <- predict(KNN_Fit1, newdata = validationSet)

confusionMatrix(KNN_Class1, validationSet$Dummy_AdmittedYes,
                positive = '1')$overall[1]


myData1$Dummy_GenderF <- as.factor(myData1$Dummy_GenderF)
myData1$Dummy_AdmittedYes <- as.factor(myData1$Dummy_AdmittedYes)

preProcessValue <- preProcess(myData1[ , 1:6], method = c("center", "scale")) 
myDataScore1 <- predict (preProcessValue, myDataScore.st)
myDataScore1$EnrollPred <- predict(KNN_Fit1, myDataScore1)

myDataScore2 <- data.frame(myDataScore, myDataScore1$EnrollPred)
View(myDataScore2)


# Output the results to a spreadsheet

library(writexl)
write_xlsx(myDataScore2,"results_case5.xlsx")
