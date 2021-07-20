library("readxl")
library(ggplot2)
library(tidyverse)
setwd(getwd())
rm(list = ls())
#
#User can change these values
StartYear <- 2018
StartProjectionYear <- 2021
EndProjectionYear <- 2050
FileName <- 'Model Inputs.xlsx'
#
#Reading Input File
user_inputs_numeric <- read_excel(FileName, sheet = 'Numeric Inputs')
user_inputs_character <- read_excel(FileName, sheet = 'Character Inputs')
PayrollOther_Data <- read_excel(FileName, sheet = 'Payroll and Other')
Benefits_Data <- read_excel(FileName, sheet = 'Benefits')
Historical_Data <- read_excel(FileName, sheet = 'Historical Data')
Scenario_Data <- read_excel(FileName, sheet = 'Inv_Returns')
OABEAAB_Balance <- read_excel(FileName, sheet = 'OABEAAB Balance')
OABEAAB_Amortization <- read_excel(FileName, sheet = 'OABEAAB Amortization')
#
##################################################################################################################################################################
#
#Functions for later use
#Function for Present Value for Amortization
PresentValue = function(rate, nper, pmt) {
  PV = pmt * (1 - (1 + rate) ^ (-nper)) / rate * (1 + rate)
  return(PV)
}

NPV = function(rate, cashflows) {
  for(i in 1:length(cashflows)){
    if(i == 1){
      NPV <- cashflows[i]/((1+rate)^(i))
    } else {
      NPV <- NPV + cashflows[i]/((1+rate)^(i))
    }
  }
  
  return(NPV)
}

#Function for calculating amo payments
#pmt0 = basic amo payment calculation, assuming payment beginning of period 
PMT0 <- function(r, nper, pv) {
  if (r == 0) {
    a <- pv/nper
  } else {
    a <- pv*r*(1+r)^(nper-1)/((1+r)^nper-1)  
  }
  
  if(nper == 0){
    a <- 0
  }
  
  return(a)
}

#pmt = amo payment function with growth rate and timing added; t = 1 for end of period payment, 0.5 for half period. 
PMT <- function(r, g = 0, nper, pv, t = 1) {
  a <- PMT0((1+r)/(1+g) - 1, nper, pv*(1+r)^t)
  return(a)
}
#
##################################################################################################################################################################
#
#Reading Values from Input input file and assigning values
#Assigning numeric inputs
for(i in 1:nrow(user_inputs_numeric)){
  if(!is.na(user_inputs_numeric[i,2])){
    assign(as.character(user_inputs_numeric[i,2]),as.double(user_inputs_numeric[i,3]))
  }
}
#
#Assigning character inputs
for(i in 1:nrow(user_inputs_character)){
  if(!is.na(user_inputs_character[i,2])){
    assign(as.character(user_inputs_character[i,2]),as.character(user_inputs_character[i,3]))
  }
}

#Create an empty Matrix for the Projection Years
EmptyMatrix <- matrix(0,(EndProjectionYear - StartProjectionYear + 1), 1)
for(j in 1:length(colnames(Historical_Data))){
  TempMatrix <- rbind(as.matrix(Historical_Data[,j]), EmptyMatrix)
  assign(as.character(colnames(Historical_Data)[j]), TempMatrix)
}
#Assign values for Projection Years
FYE <- StartYear:EndProjectionYear
#Get Start Index, since historical data has 3 rows, we want to start at 4
StartIndex <- StartProjectionYear - StartYear + 1
HistoricalIndex <- StartProjectionYear - StartYear

#Initialize the Inflation Adjusted Variables to use later
#UAL_AVA_InflAdj <- UAL_AVA
#colnames(UAL_AVA_InflAdj) <- 'UAL_AVA_InflAdj'
#UAL_MVA_InflAdj <- UAL_MVA
#colnames(UAL_MVA_InflAdj) <- 'UAL_MVA_InflAdj'

#Assign values for simulation
if(SimType == 'Assumed'){
  SimReturn <- SimReturnAssumed
} else if(AnalysisType == 'Conservative'){
  SimReturn <- SimReturnConservative
}

#Initialize Amortization and Outstnading Base
#Regular Amortization
RowColCount_Regular <- (EndProjectionYear - StartProjectionYear + 1)
OutstandingBase_Regular <- matrix(0,RowColCount_Regular, RowColCount_Regular + 1)
Amortization_Regular <- matrix(0,RowColCount_Regular, RowColCount_Regular + 1)
AmoYearsInput_Regular <- read_excel(FileName, sheet = 'Amortization')$'Amortization - Regular'

#Experience Account
RowColCount_ExpAccount <- (EndProjectionYear - StartProjectionYear + 1)
OutstandingBase_ExpAccount <- matrix(0,RowColCount_ExpAccount, RowColCount_ExpAccount + 1)
Amortization_ExpAccount <- matrix(0,RowColCount_ExpAccount, RowColCount_ExpAccount + 1)
AmoYearsInput_ExpAccount <- read_excel(FileName, sheet = 'Amortization')$'Amortization - Experience Account'

#Contribution Variance
RowColCount <- (EndProjectionYear - StartProjectionYear + 1)
OutstandingBase_ContrVar <- matrix(0,RowColCount, 5)
Amortization_ContrVar <- matrix(0,RowColCount, 5)
#
##################################################################################################################################################################
#
#Offset Matrix. Used later for Amortization calculation
#Regular Amortization
OffsetYears_Regular <- matrix(0,RowColCount_Regular, RowColCount_Regular)
for(i in 1:nrow(OffsetYears_Regular)){
  RowCount_Regular <- i
  ColCount_Regular <- 1
  #This is to create the "zig-zag" pattern of amo years. you also want it to become 0 if the amo years is greater than the payments
  #Meaning if its 10 years, then after 10 years, the amo payment is 0
  while(RowCount_Regular <= nrow(OffsetYears_Regular) && (ColCount_Regular <= as.double(AmoYearsInput_Regular[i]))){
    OffsetYears_Regular[RowCount_Regular,ColCount_Regular] <- as.double(AmoYearsInput_Regular[i])
    RowCount_Regular <- RowCount_Regular + 1
    ColCount_Regular <- ColCount_Regular + 1
  }
}

for(i in 1:nrow(OffsetYears_Regular)-1){
  for(j in 1:i+1){
    if(OffsetYears_Regular[i+1,j] > 0) {
      OffsetYears_Regular[i+1,j] <- max(OffsetYears_Regular[i+1,j] - j + 1, 1) 
    }
  }
}

#Experience Account
OffsetYears_ExpAccount <- matrix(0,RowColCount_ExpAccount, RowColCount_ExpAccount)
for(i in 1:nrow(OffsetYears_ExpAccount)){
  RowCount_ExpAccount <- i
  ColCount_ExpAccount <- 1
  #This is to create the "zig-zag" pattern of amo years. you also want it to become 0 if the amo years is greater than the payments
  #Meaning if its 10 years, then after 10 years, the amo payment is 0
  while(RowCount_ExpAccount <= nrow(OffsetYears_ExpAccount) && (ColCount_ExpAccount <= as.double(AmoYearsInput_ExpAccount[i]))){
    OffsetYears_ExpAccount[RowCount_ExpAccount,ColCount_ExpAccount] <- as.double(AmoYearsInput_ExpAccount[i])
    RowCount_ExpAccount <- RowCount_ExpAccount + 1
    ColCount_ExpAccount <- ColCount_ExpAccount + 1
  }
}

for(i in 1:nrow(OffsetYears_ExpAccount)-1){
  for(j in 1:i+1){
    if(OffsetYears_ExpAccount[i+1,j] > 0) {
      OffsetYears_ExpAccount[i+1,j] <- max(OffsetYears_ExpAccount[i+1,j] - j + 1, 1) 
    }
  }
}

#Initialize Outstanding Balances
OutstandingBase_ContrVar[1,1] <- VarCont_ARC_ERCont[HistoricalIndex] + 16.668183
if(FR_NewDR[HistoricalIndex] >= ResetBases){
  OutstandingBase_ExpAccount[1,1] <- 0
  OutstandingBase_Regular[1,1] <- 0
} else if(FR_NewDR[HistoricalIndex] < 0.85){
  OutstandingBase_ExpAccount[1,1] <- FinalAllocExcRet[HistoricalIndex]
} else {
  OutstandingBase_ExpAccount[1,1] <- GainAllocOAB[HistoricalIndex] + FinalAllocExcRet[HistoricalIndex] - ActGainLoss[HistoricalIndex]
  OutstandingBase_Regular[1,1] <- UAL_NewDR[HistoricalIndex] - (OABBalance[HistoricalIndex] + EAABBalance[HistoricalIndex] + OtherBalance[HistoricalIndex] + VarContrBalance[HistoricalIndex]) - OutstandingBase_ExpAccount[1]
}

#Initialize Amortization Values
Amortization_ContrVar[1,1] <- -OutstandingBase_ContrVar[1,1] / PresentValue(NewDR[HistoricalIndex],5,1)
Amortization_ExpAccount[1,1] <- -OutstandingBase_ExpAccount[1,1] / PresentValue(NewDR[HistoricalIndex],OffsetYears_ExpAccount[1],1)
Amortization_Regular[1,1] <- -OutstandingBase_Regular[1,1] / PresentValue(NewDR[HistoricalIndex],OffsetYears_Regular[1],1)
#
##################################################################################################################################################################
#
#This is payroll upto acrrued liability. these values do not change regardless of scenario or other stress testing
#In order to optimize, they are placed outside the function
for(i in StartIndex:length(FYE)){
  #Payroll
  if(PayrollAssumption == 'Assumed'){
    TotalPayroll[i] <- TotalPayroll[i-1]*(1 + Payroll_growth) 
  } else{
    TotalPayroll[i] <- TotalPayroll[i-1]*(1 + PayrollOther_Data$`Smoothed Payroll Growth`[i])
  }
  ClosedPlans[i] <- ClosedPlans[i-1]*PayrollOther_Data$`Closed Plans`[i]/PayrollOther_Data$`Closed Plans`[i-1]
  NewHirePayroll[i] <- (TotalPayroll[i] - ClosedPlans[i])*PayrollOther_Data$`New Tier`[i]
  LegacyPayroll[i] <- TotalPayroll[i] - (ClosedPlans[i] + NewHirePayroll[i])
  ARCPayroll[i] <- TotalPayroll[i-1]*(1 + Payroll_growth)^0.5
  
  #Benefits, Refunds, Admin
  OrigBP[i] <- -Benefits_Data$`ORIG NO COLA BP`[i]
  TotalBP[i] <- -Benefits_Data$`Total BP`[i]
  COLA_CumBP[i] <- TotalBP[i] - OrigBP[i]
  Admin[i] <- Admin_Exp_Pct*TotalPayroll[i]
  Refunds[i] <- Refunds_Pct*TotalPayroll[i]
  
  #Original DR for NC, Liabilities, Benefits
  OriginalDR[i] <- dis_r_proj
  NewDR[i] <- dis_r_proj
  MOYNC_OrigDR[i] <- ClosedPlans[i-1]*GrossNC_ClosedPlans + NewHirePayroll[i-1]*GrossNC_NewHires + LegacyPayroll[i-1]*GrossNC_OldHires
  BOYAccrLiabOrigDR_Total[i] <- (MOYNC_OrigDR[i] + BOYAccrLiabOrigDR_Total[i-1])*(1 + OriginalDR[i-1]) + OrigBP[i]*(1 + OriginalDR[i-1])^0.5
  
  #COLAPV
  #COLA PV has to be calcd for new DR before old DR
  if((InclPlan == 'Yes') && (FYE[i] == 2021)){
    COLAPV[i] <- -NPVCOLA
  } else {
    COLAPV[i] <- 0
  }
  COLA_PV_NewDR[i] <- max(COLA_PV_NewDR[i-1]*(1+NewDR[i-1]) + COLA_CumBP[i]*(1+NewDR[i-1])^0.5 - COLAPV[i], 0)
  COLA_PV_OrigDR[i] <- COLA_PV_NewDR[i]
  EOYAccrLiabOrigDR_Total[i] <- BOYAccrLiabOrigDR_Total[i] + COLA_PV_OrigDR[i]
  
  #New DR calc for MOY NC and Liabilities
  DRDifference <- 100*(OriginalDR[i] - NewDR[i])
  MOYNC_NewDR[i] <- MOYNC_OrigDR[i]*((1+(NCSensDR/100))^(DRDifference))
  BOYAccrLiabNewDR_Total[i] <- BOYAccrLiabOrigDR_Total[i]*((1+(LiabSensDR/100))^(DRDifference))*((1+(Convexity/100))^((DRDifference)^2/2))
  EOYAccrLiabNewDR_Total[i] <- EOYAccrLiabOrigDR_Total[i]*((1+(LiabSensDR/100))^(DRDifference))*((1+(Convexity/100))^((DRDifference)^2/2))
  
  #Contribution Rates
  EEContrib[i] <- (ClosedPlans[i]*ClosedPlanEEConrib + NewHirePayroll[i]*EEContrib_NewHires + LegacyPayroll[i]*EEContrib_OldHires)*(ARCPayroll[i]/TotalPayroll[i])
  EE_Contrib_Pct[i] <- (ClosedPlans[i]*ClosedPlanEEConrib + NewHirePayroll[i]*EEContrib_NewHires + LegacyPayroll[i]*EEContrib_OldHires) / TotalPayroll[i]
  ER_NC[i] <- MOYNC_NewDR[i]*(1 + NewDR[i-1])^0.5 - EEContrib[i]
  GrossTotalNC_Pct[i] <- (EEContrib[i] + ER_NC[i]) / TotalPayroll[i]
  ER_ARC_NC_Pct[i] <- GrossTotalNC_Pct[i] - EE_Contrib_Pct[i]
  ER_Proj_NC_Pct[i] <- ER_ARC_NC_Pct[i-1]
}
#
##################################################################################################################################################################
#Running the model from NC onwards. this gets iterated for different scenarios
#RunModel <- function(AnalysisType, SimReturn, SimVolatility, ER_Policy, ScenType, SupplContrib, LvDollarorPercent, SegalOrReason){
#Scenario Index for referencing later based on investment return data
ScenarioIndex <- which(colnames(Scenario_Data) == as.character(ScenType))

#intialize this value at 0 for Total ER Contributions
#Total_ER[StartIndex-1] <- 0
for(i in StartIndex:length(FYE)){
  print(i)
  ProjectionCount <- i - StartIndex + 1
  #OAAB, EAAB, etc. balances and Amortizations
  if(max(FR_NewDR[1:i-1]) >= ResetBases){
    #Balance
    OABBalance[i] <- 0
    EAABBalance[i] <- 0
    OtherBalance[i] <- 0
    VarContrBalance[i] <- 0
    #Amortization
    OABAmortization[i] <- 0
    EAABAmortization[i] <- 0
    OtherAmortization[i] <- 0
    VarContrAmortization[i] <- 0
  } else {
    #Balance
    OABBalance[i] <- OABEAAB_Balance$OAB[i] / 1e6
    EAABBalance[i] <- OABEAAB_Balance$EAAB[i] / 1e6
    OtherBalance[i] <- sum(OABEAAB_Balance[6:33]) / 1e6
    VarContrBalance[i] <- OABEAAB_Balance$`Cont. Var` / 1e6
    #Amortization
    OABAmortization[i] <- OABEAAB_Amortization$OAB[i-1] / 1e6
    EAABAmortization[i] <- OABEAAB_Amortization$EAAB[i-1] / 1e6
    OtherAmortization[i] <- sum(OABEAAB_Amortization[i-1,6:33]) / 1e6
    VarContrAmortization[i] <- OABEAAB_Amortization$`Cont. Var`[i-1] / 1e6
  }
  #Calculate the amo payments, layered bases, etc.
  TotalCurrAmoPmt[i] <- OABAmortization[i] + EAABAmortization[i] + OtherAmortization[i] + VarContrAmortization[i]
  LayeredBases[i] <- sum(Amortization_ContrVar[ProjectionCount,]) + sum(Amortization_ExpAccount[ProjectionCount,]) + sum(Amortization_Regular[ProjectionCount,])
  UAL_ARC_Pct[i] <- (TotalCurrAmoPmt[i] + LayeredBases[i]) / ARCPayroll[i]
  ER_ARC_Pct[i] <- max(0,UAL_ARC_Pct[i] + ER_ARC_NC_Pct[i] - 0.001)
  
  #This is the floor for the ER min UAL payment
  if(OABEAAB_Balance$OAB[i] > 0){
    ER_OAB_Floor <- StatMin_OAB
  } else {
    ER_OAB_Floor <- StatMin
  }
  if(FR_NewDR[i-1] > ContrHoliday){
    ER_Proj_UAL_Pct[i] <- 0
    ARC_MidYr_ReqCont[i] <- 0
  } else {
    ER_Proj_UAL_Pct[i] <- max((TotalCurrAmoPmt[i] + LayeredBases[i]) / TotalPayroll[i] / (1+NewDR[i-1])^0.5, ER_OAB_Floor) - ER_Proj_NC_Pct[i]
    ARC_MidYr_ReqCont[i] <- (UAL_ARC_Pct[i] + ER_ARC_NC_Pct[i])*ARCPayroll[i]
  }
  NoER_Min_Proj_UAL_Pct[i] <- (TotalCurrAmoPmt[i] + LayeredBases[i]) / TotalPayroll[i] / (1+NewDR[i-1])^0.5 - ER_Proj_NC_Pct[i]
  ER_Contrib_Pct[i] <- ER_Proj_UAL_Pct[i] + ER_Proj_NC_Pct[i]
  
  #Calc for total contribution rate
  if(FR_NewDR[i-1] <= ContrHoliday){
    UAL_Payment[i] <- max(NoER_Min_Proj_UAL_Pct[i]*ARCPayroll[i], -ER_NC[i])
  } else if((NoER_Min_Proj_UAL_Pct[i]*ARCPayroll[i] < 0) | (NoER_Min_Proj_UAL_Pct[i]*ARCPayroll[i]) < -ER_NC[i]){
    UAL_Payment[i] <- -ER_NC[i]
  } else {
    UAL_Payment[i] <- NoER_Min_Proj_UAL_Pct[i]*ARCPayroll[i]
  }
  ER_Min[i] <- (ER_Proj_UAL_Pct[i] - NoER_Min_Proj_UAL_Pct[i])*ARCPayroll[i]
  TotalContrib[i] <- ER_Min[i] + UAL_Payment[i] + ER_NC[i]
                                                         
  #Return data based on deterministic or stochastic
  if((AnalysisType == 'Stochastic') && (i >= StartIndex)){
    ROA_MVA[i] <- rnorm(1,SimReturn,SimVolatility)
  } else if(AnalysisType == 'Deterministic'){
    ROA_MVA[i] <- as.double(Scenario_Data[i,ScenarioIndex]) 
  }
  
  #Cashflows and total defered
  NetCF[i] <- TotalBP[i] + Refunds[i] + Admin[i] + EEContrib[i] + TotalContrib[i]
  MVA[i] <- MVA[i-1]*(1+ROA_MVA[i]) + NetCF[i]*(1+ROA_MVA[i])^0.5
  ActInvReturn[i] <- MVA[i] - MVA[i-1] - NetCF[i]
  ExpReturn[i] <- MVA[i-1]*NewDR[i-1] + (NetCF[i] - LegisContrib[i])*NewDR[i-1]*0.5
  GainLoss_CurrentYear[i] <- ActInvReturn[i] - ExpReturn[i]
  DeferedCurYearGL[i] <- GainLoss_CurrentYear[i]*(4/5)
  Year1GL[i] <- DeferedCurYearGL[i]*(3/4)
  Year2GL[i] <- Year1GL[i]*(2/3)
  Year3GL[i] <- Year2GL[i]*(1/2)
  TotalDefered[i] <- DeferedCurYearGL[i] + Year1GL[i] + Year2GL[i] + Year3GL[i]
  
  #AVA and actuarial returns
  AVA[i] <- min((1+AVA_Corridor)*MVA[i],max(MVA[i]-TotalDefered[i],(1-AVA_Corridor)*MVA[i]))
  ActuarialActRet[i] <- AVA[i] - AVA[i-1] - NetCF[i]
  ActInvRet_Pct[i] <- 2*ActuarialActRet[i] / (AVA[i] + AVA[i-1] - ActuarialActRet[i])
  
  ActExpRet[i] <- AVA[i-1]*NewDR[i-1] + NetCF[i]*NewDR[i-1]*0.5
  ActGainLoss[i] <- ActuarialActRet[i] - ActExpRet[i]
  Threshold[i] <- Threshold[i-1]*(AVA[i]/AVA[i-1])
  RemainingOAB[i] <- (OABEAAB_Balance$`Prelim OAB`[i] + OABEAAB_Balance$`Prelim EAAB`[i])/1e6
  GainAllocOAB[i] <- min(max(0,ActGainLoss[i]),RemainingOAB[i],Threshold[i])
  GainAllocExpAcct[i] <- 0.5*max(ActGainLoss[i] - GainAllocOAB[i],0)
  IntAllocExpAcct[i] <- ActuarialActRet[i]*ExpAcc[i-1] / AVA[i-1]
  PrelimEOYBal[i] <- ExpAcc[i-1] + IntAllocExpAcct[i] + GainAllocExpAcct[i]
  
  if(FR_NewDR[i-1] < 0.8){
    AccountCap[i] <- Benefits_Data$`PV of PBI`[i]
  } else {
    AccountCap[i] <- 2*Benefits_Data$`PV of PBI`[i]
  }
  ExpAcc[i] <- min(AccountCap[i], COLAPV[i] + PrelimEOYBal[i])
  
  if(ActInvRet_Pct[i] < 0.0825){
    PBI_Pct_1 = 0.02
  } else {
    PBI_Pct_1 = 9.99
  }
  if(FR_NewDR[i-1] >= 0.8){
    PBI_Pct_2 = 0.03
  } else if(FR_NewDR[i-1] >= 0.75){
    PBI_Pct_2 = 0.025
  } else if(FR_NewDR[i-1] >= 0.65){
    PBI_Pct_2 = 0.02
  } else if(FR_NewDR[i-1] >= 0.55){
    PBI_Pct_2 = 0.015
  } else if(FR_NewDR[i-1] < 0.55){
    PBI_Pct_2 = 0
  }
  PBIPercent <- min(PBI_Pct_1,PBI_Pct_2)
  FinalAllocExcRet[i] <- ExpAcc[i] - ExpAcc[i-1] - IntAllocExpAcct[i] - COLAPV[i]
  
  ER_Interest[i] <- ER_CreditAccount[i-1] / ActInvRet_Pct[i]
  ER_Inflow[i] <- max(0,ER_Proj_UAL_Pct[i] - NoER_Min_Proj_UAL_Pct[i])*ARCPayroll[i]
  VarCont_ARC_ERCont[i] <- max(ARC_MidYr_ReqCont[i] - TotalContrib[i] + ER_Inflow[i],0)
  TempValue <- EOYAccrLiabNewDR_Total[i] + MVA[i]+TotalDefered[i] + ExpAcc[i] + ER_Interest[i]
  ER_Outflow[i] <- -min(max(0,TempValue + ER_CreditAccount[i-1] + ER_Inflow[i]), ER_CreditAccount[i-1] + ER_Inflow[i])
  ER_CreditAccount[i] <- ER_CreditAccount[i-1] + ER_Interest[i] + ER_Inflow[i] + ER_Outflow[i]
  EOYValAssets[i] <- MVA[i] - TotalDefered[i] - ExpAcc[i] - ER_CreditAccount[i]
  
  FR_NewDR[i] <- EOYValAssets[i] / EOYAccrLiabNewDR_Total[i]
  UAL_NewDR[i] <- EOYAccrLiabNewDR_Total[i] - EOYValAssets[i]
  
  if(FR_NewDR[i] >= ResetBases){
    OutstandingBase_ContrVar[ProjectionCount+1,] <- 0
    OutstandingBase_ExpAccount[ProjectionCount+1,] <- 0
    OutstandingBase_Regular[ProjectionCount+1,] <- 0
  } else if(FR_NewDR[i] < 0.85){
    OutstandingBase_ExpAccount[ProjectionCount+1,2:(ProjectionCount + 1)] <- OutstandingBase_ExpAccount[ProjectionCount,1:ProjectionCount]*(1 + NewDR[i-1]) - (Amortization_ExpAccount[ProjectionCount,1:ProjectionCount]*(1 + NewDR[i-1])^0.5)
    OutstandingBase_ExpAccount[ProjectionCount+1,1] <- FinalAllocExcRet[i]
  } else {
    OutstandingBase_ContrVar[ProjectionCount+1,2:5] <- OutstandingBase_ContrVar[ProjectionCount,1:4]*(1 + NewDR[i-1]) - (Amortization_ContrVar[ProjectionCount,1:4]*(1 + NewDR[i-1])^0.5)
    OutstandingBase_ContrVar[ProjectionCount+1,1] <- VarCont_ARC_ERCont[i]
    
    OutstandingBase_ExpAccount[ProjectionCount+1,2:(ProjectionCount + 1)] <- OutstandingBase_ExpAccount[ProjectionCount,1:ProjectionCount]*(1 + NewDR[i-1]) - (Amortization_ExpAccount[ProjectionCount,1:ProjectionCount]*(1 + NewDR[i-1])^0.5)
    OutstandingBase_ExpAccount[ProjectionCount+1,1] <- GainAllocOAB[i] + FinalAllocExcRet[i] - ActGainLoss[i]
    
    OutstandingBase_Regular[ProjectionCount+1,2:(ProjectionCount + 1)] <- OutstandingBase_Regular[ProjectionCount,1:ProjectionCount]*(1 + NewDR[i-1]) - (Amortization_Regular[ProjectionCount,1:ProjectionCount]*(1 + NewDR[i-1])^0.5)
    OutstandingBase_Regular[ProjectionCount+1,1] <- UAL_NewDR[i] - (OABBalance[i] + EAABBalance[i] + OtherBalance[i] + VarContrBalance[i]) - OutstandingBase_ExpAccount[ProjectionCount,1]
  }
  
  #Initialize Amortization Values
  
  Amortization_ContrVar[ProjectionCount+1,1:5] <- PMT(pv = OutstandingBase_ContrVar[ProjectionCount+1,1:5], 
                                                                              r = NewDR[i-1], g = AmoBaseInc, t = 0.5, nper = 5)
  
  Amortization_ExpAccount[ProjectionCount+1,1:(ProjectionCount + 1)] <- PMT(pv = OutstandingBase_ExpAccount[ProjectionCount+1,1:(ProjectionCount + 1)], 
                                                                              r = NewDR[i-1], g = AmoBaseInc, t = 0.5,
                                                                              nper = pmax(OffsetYears_ExpAccount[ProjectionCount+1,1:(ProjectionCount + 1)]))
  
  Amortization_Regular[ProjectionCount+1,1:(ProjectionCount + 1)] <- PMT(pv = OutstandingBase_Regular[ProjectionCount+1,1:(ProjectionCount + 1)], 
                                                                            r = NewDR[i-1], g = AmoBaseInc, t = 0.5,
                                                                            nper = pmax(OffsetYears_Regular[ProjectionCount+1,1:(ProjectionCount + 1)]))
}

#return(Output)
#}
#
# ##################################################################################################################################################################
#
