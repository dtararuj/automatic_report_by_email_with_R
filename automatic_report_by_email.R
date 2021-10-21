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
OutMail[["Subject"]]="X report"
OutMail[["body"]]="report attached below"

#indicate the attachment, our report
OutMail[["Attachments"]]$Add(file.path(getwd(),"raporttest.csv"))

#3 Send an email when certain condition will be met 
ifelse(report %>%  select(2) %>%  sum() > 499,OutMail$Send(),"email not sent")


#4 Send an email to defined recipients

recipients = c("dawidtararuj@10g.pl", "dawidtataruj@10g.pl","dawidtararuj2@10g.pl")

if(report %>%  select(2) %>%  sum() > 499){
  for (email in recipients){
    OutApp<-COMCreate("Outlook.Application")
    OutMail=OutApp$CreateItem(0)
    OutMail[["TO"]]= email
    OutMail[["Subject"]]="X report"
    OutMail[["body"]]="Report attached below"
    OutMail[["Attachments"]]$Add("C:/Users/Dawid/Desktop/raporttest.csv")
    OutMail$Send()}
  }else{
    stop()
  }


