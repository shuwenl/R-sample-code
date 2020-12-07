

source ("D:/_Programs/R/Analysis_Functions.r")

#connect and explore data
A<-data_use(dbname="sample.sqlite")


#create a new variable and test correlation
B<-corr_Calcu(A)

#create pictures to explore data
Picture(B)

#Statistical Analysis
#T test, lm
Stat(B)

#logistic reg
logis(B)






















