##Script use to send e-mail with report, if one condition will be met.

#install.packages("RDCOMClient", repos="http://www.omegahat.net/R")

library(tidyverse)
library(RDCOMClient)


# 1. Prepare a sample report 
report = iris %>% group_by(Species) %>%  summarise_all(.funs = sum)

write.csv(report, "raporttest.csv")

#2. E-mail preparation
# a) in some case you have to first open your outlook app manually

OutApp<-COMCreate("Outlook.Application")
OutMail=OutApp$CreateItem(0)

#define recipient, subject and message text
OutMail[["TO"]]= "dawidtararuj@10g.pl"
OutMail[["Subject"]]="raport3"
OutMail[["body"]]="raport w za³aczniku"

#indicate the attachment, our report
OutMail[["Attachments"]]$Add(file.path(getwd(),"raporttest.csv"))

#3 Send an email when certain condition will be met 
ifelse(report %>%  select(2) %>%  sum() > 499,OutMail$Send(),"email not sent")

