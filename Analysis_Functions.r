#load libraries and data
#library (data.table)
#library (plyr)
#library (stringr)
#library(matlib)
#library('xlsx')
#library(e1071) 
#library(gplots)
library("ggpubr")
#library(multcomp)
#library(car)
#library(aod)
#library(ggplot2)
#library("dplyr")
#install.packages('RSQLite')
library(rstudioapi)    
library(RSQLite)
#install.packages('DBI')
#library(DBI)
#install.packages('tidyverse')
library(tidyverse)
#install.packages('odbc')
#library(odbc)
#connect to SQLite
library(officer)
library("ggplotify")


data_use<-function(dbname="")
{
  con.sqlite <-  DBI::dbConnect(RSQLite::SQLite(), dbname=dbname)
  df0<-dbGetQuery(con.sqlite, "select * from new;")

  #check the data
  set.seed(1234)
  dplyr::sample_n(df0, 10)

  #check structure
  str(df0)
  #print(df0)

  #Data exploration
  summary(df0)
  mean(df0$IPCNT)
  attach(df0)
  median(IPDays)
  
  save_chart_2_pptx(as.ggplot(function() hist(df0$TotCost)))
  save_chart_2_pptx(as.ggplot(function() hist(log(TotCost))))
  save_chart_2_pptx(as.ggplot(function() boxplot(OPCNT ~Age)))
  save_chart_2_pptx(as.ggplot(function() boxplot(OPCNT ~Race)))
  save_chart_2_pptx(as.ggplot(function() plot(TotCost ~MedCost)))#medCost is Medicine Cost
  return(df0)
}

#export to powerpoint
save_chart_2_pptx <- function(
  gg_plot, pptx_file="temp.pptx", 
  full_size = F, slide_editable = F,
  left = 0.5, top = 0.5, width = 9, height = 7){
  
  if (is.na(gg_plot)) return
  if (is.na(pptx_file)) return
  
  # if (slide_editable)
  #   gg_plot <- dml(ggobj = gg_plot)
  
  if (file.exists(pptx_file))
    doc <- read_pptx(pptx_file)
  else
    doc <- read_pptx()
  
  doc <- add_slide(doc)
  if (full_size)
    doc <- ph_with(x = doc, value = gg_plot,
                   location = ph_location_fullsize(),
                   bg = "transparent")
  else
    doc <- ph_with(x = doc, value = gg_plot,
                   location = ph_location(left = left, top = top, 
                                          width = width, height = height),
                   bg = "transparent")
  print(doc, target = pptx_file)
}

#B<-corr_Calcu(A)

corr_Calcu<-function(df01)
{
  #create a new variable
  df01<-df01%>%
    mutate(MedicalCost = TotCost-MedCost)
  save_chart_2_pptx(as.ggplot(function() plot(df01$TotCost ~df01$MedicalCost)))
#test correlation
#Kendall rank correlation test
  cor.test(df01$MedicalCost, df01$TotCost,  method="kendall")

#Spearman rank correlation coefficient
  cor.test(df01$MedicalCost, df01$TotCost,  method = "spearman")
  return(df01)
}

Picture<-function(df01)
{
 p1<- df01 %>% 
  select(ID, EDCNT, IPCNT,OPCNT,ReadmitCNT) %>% 
  pivot_longer(cols=c("EDCNT","IPCNT","OPCNT","ReadmitCNT")) %>%
  ggplot(aes(x=name,y=value)) +
  geom_col(colour='Green')
 save_chart_2_pptx (p1)
 
 p2<-df01 %>% 
  select(ID, MedicalCost, MedCost,TotCost) %>% 
  pivot_longer(cols=c("MedicalCost","MedCost","TotCost")) %>%
  ggplot(aes(x=name,y=value)) +
  geom_col(colour='Red')
 save_chart_2_pptx(p2)
 
 p3<-df01%>% 
  select(ID,Age,EDCNT,IPCNT,IPDays,OPCNT,ReadmitCNT,
         TotCost,MedCost,MedicalCost) %>% 
  pivot_longer(cols=c("Age","EDCNT","IPCNT","IPDays","OPCNT",
                      "ReadmitCNT","TotCost","MedCost","MedicalCost")) %>%
  ggplot(aes(sample = value)) +
  stat_qq(aes(,colour=name)) +
  stat_qq_line(colour="black") +
  facet_wrap(~ name, scales = "free")
 save_chart_2_pptx(p3)
#Boxplot
 p4<-df01 %>% 
  filter(Sex == 'F') %>%
  ggplot(aes(x=Race, y=IPCNT)) +
  geom_boxplot(colour='Blue',fill='yellow')
 save_chart_2_pptx(p4)
 
 p5<-df01 %>% 
  filter(Sex == 'M',Race !='NA') %>%
  ggplot(aes(x=Race, y=OPCNT)) +
  geom_boxplot(colour="orange", fill="yellow")
 save_chart_2_pptx(p5)

 p6<-df01%>%
  select(TotCost, Age,Race)%>%
  filter( Age>30,Age<50 )%>%
  group_by(Race)%>%
  summarise(mean(TotCost))
  save_chart_2_pptx(p6)
 
  p7<-df01%>%
  select(TotCost, Race,Age)%>%
  filter(( Race == "White"|Race == "BLACK") & Age<50)%>%
  ggplot(aes(x=Race, y=TotCost)) +
  geom_boxplot(colour="black", fill="red")
  save_chart_2_pptx(p7)
return(df01)
}

#Statistical Analysis
Stat<-function(df01)
{
  #T test
  df03<-df01%>%
  select(TotCost, Race,Age)%>%
  filter(( Race == "White"|Race == "BLACK") & Age<50)

 t.test (data =df03, TotCost ~Race)
 wilcox.test(TotCost ~ Race, data = df03,exact = FALSE) #non-parametric 

#linear regression
 p8<-df01%>%
  filter(Age<50)%>%
  ggplot(aes(x=OPCNT, y=TotCost,fill=Sex,colour=Sex))+
  geom_point(alpha=0.8)+
  geom_smooth(method=lm)+
  facet_wrap(~Race)
  save_chart_2_pptx(p8)
#Self-Exploring
#check row # and distinct values
#df01%>% nrow()
#df01$OPCNT %>% n_distinct()

 #check distribution
 KS<-ks.test(df01$OPCNT, "pnorm", mean=mean(df01$OPCNT), sd=sd(df01$OPCNT))
 class(df01$OPCNT)

 #SW<-shapiro.test(df01$OPCNT)#sample size must between 3 and 5000

 summary(lm(df01$TotCost ~df01$OPCNT))

 #TotCost vs MedicalCost
 Single<-df01%>%
  select(TotCost, MedicalCost)
#check the data
 summary(Single)
#remove outliers
#install.packages("ggstatsplot")
# Load the package
 library(ggstatsplot)
#check outliers
 #boxplot(Single)$out
#add one field RiskFlag
}

logis<-function(df01)
{
  df04<-df01%>%
  mutate(RiskFlag = ifelse(OPCNT %in% 0:50,"0",
                              ifelse(OPCNT %in% 50:246,"1",NA)))

#check for missing values and look how many unique values 
#there are for each variable
 sapply(df04,function(x) sum(is.na(x)))
 sapply(df04, function(x) length(unique(x)))
 print('stop1')
 str(df04)
 
#selecting the relevant columns of original dataset
 df<- subset(df04,select=c(2,3,4,5,9,10,13))
 
#A factor is how R deals categorical variables
 df$Sex<-factor(df$Sex)
 
#split the data into two part
 train <- df[1:9900,]
 test <- df[9901:10000,]

#fit the model
 train$RiskFlag <- as.numeric(as.character(train$RiskFlag))

 model <- glm(RiskFlag ~ Age + Sex + EDCNT + IPCNT + ReadmitCNT + TotCost,family=binomial(link='logit'),data=train)
 summary(model)

#run anova() function on the model to analyze the table of deviance
 anova(model, test="Chisq")

#McFadden R2 index is used to assess the model fit
 install.packages('pscl')
 library(pscl)
 pR2(model)
print('stop2')
#Assessing the predictive ability of the model
 fitted.results <- predict(model,newdata=subset(test,select=c(1,2,3,4,5,6)),type='response')
 fitted.results <- ifelse(fitted.results > 0.5,1,0)
 misClasificError <- mean(fitted.results != test$RiskFlag)
 print(paste('Accuracy',1-misClasificError))

 #Calculate AUC
 install.packages('ROCR')
 library(ROCR)
 p <- predict(model, newdata=subset(test,select=c(1,2,3,4,5,6)), type="response")
 pr <- prediction(p, test$RiskFlag)
 #plot the ROC curve and calculate the AUC (area under the curve) 
 prf <- performance(pr, measure = "tpr", x.measure = "fpr")
 plot(prf)

 auc <- performance(pr, measure = "auc")
 auc <- auc@y.values[[1]]
 auc
 return(df04)
}





















