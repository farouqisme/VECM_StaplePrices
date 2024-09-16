rm(list=ls())

setwd("D:/KeRjAaaAAAaaAaAAA/paper hargapangan")

library(devtools)
source_url("https://raw.githubusercontent.com/anguyen1210/var-tools/master/R/extract_varirf.R")



library(tidyverse)
library(tseries)
library(dplyr)
library(rio)
library(ggplot2)
library(scales)
library(vars)
library(urca)
library(stringr)
library(gridExtra)
library(lubridate)
library(tsDyn)   # VECM Estimation
library(officer)
###take .xlsx name in a folder
listfile <- list.files(pattern = "Indonesia")

###import .xlsx based on list of names and append in a list
dat <- list()
for(i in 1:length(listfile)){
  a <- readxl::read_excel(listfile[i],sheet = 2)
  dat <- append(dat,list(a))
} 



###using regex to clean names
datname <- listfile %>% gsub("^Indonesia |\\.xlsx$", "",.) %>% tolower()
names(dat) <- datname

list2env(dat, envir = .GlobalEnv)

###calculate rowMeans and subset variable then assign into 
##new data.frame with the same name
for(i in 1:length(datname)){
  nama <- datname[i]
  a <- get(datname[i])
  a[paste(nama)] <- rowMeans(get(datname[i])[,c(2:88)])
  a <- a %>% dplyr::select(.,Date,paste(nama))
  a$Date <- as.Date(a$Date)
  assign(nama, a, envir = .GlobalEnv)
}

##merge each data.frames
datlist <- lapply(datname, get)
comp_dat2 <- Reduce(function(x, y) merge(x, y, by = "Date", all = TRUE), datlist)
comp_dat <- list()
for(i in 2:ncol(comp_dat2)){
  a <- log(comp_dat2[i])
  comp_dat <- append(comp_dat,a)
}
comp_dat <- data.frame(comp_dat)
comp_dat$Date <- comp_dat2$Date
comp_dat <- comp_dat[, c("Date", names(comp_dat)[names(comp_dat) != "Date"])]


######## Remove Saturday & Sunday ##########
comp_dat_level <- comp_dat
comp_dat_level$day <- weekdays(comp_dat_level$Date)
comp_dat_level <- comp_dat_level[comp_dat_level$day != "Saturday",]
comp_dat_level <- comp_dat_level[comp_dat_level$day != "Sunday",]
comp_dat_level <- comp_dat_level %>% dplyr::select(., c(-day))
# openxlsx::write.xlsx(comp_dat_level,"data level.xlsx")
comp_dat_level <- comp_dat_level %>% dplyr::select(., c(-Date))

comp_dat <- comp_dat_level    ######<-The estimation starts from log transformation data



######## UNIT ROOT ##########

##adf.test
dtset2 <- list()
adf <- list()
pp <- list()
adf_d1 <- list()
pp_d1 <- list()
for(i in 1:ncol(comp_dat)){
  dtset2[[i]]<-ts(comp_dat[,i],start = c(year = 2018, month = 1, day = 1),frequency = 366)
  a <- adf.test(dtset2[[i]])
  b <- adf.test(diff(dtset2[[i]]))
  
  c <- pp.test(dtset2[[i]])
  d <- pp.test(diff(dtset2[[i]]))
  
  adf <- append(adf,a["p.value"])
  adf_d1 <- append(adf_d1,b["p.value"])
  pp <- append(pp,c["p.value"])
  pp_d1 <- append(pp_d1,d["p.value"])
  names(dtset2)[i] <- colnames(comp_dat)[i]
}

for(i in 1:length(adf)){
  names(adf)[i] <- paste0(datname[i],"_adf")
  names(adf_d1)[i] <- paste0(datname[i],"_adf")
  names(pp)[i] <- paste0(datname[i],"_pp")
  names(pp_d1)[i] <- paste0(datname[i],"_pp")
}
stat_test <- list("adf" = data.frame(adf),"adf_d1" = data.frame(adf_d1),
                  "pp" = data.frame(pp),"pp_d1" = data.frame(pp_d1))

stat_test2 <- t(data.frame(stat_test))
stat_test_df <- data.frame(stat_test) ### <-Stationarity test result (FIRST DIFF)
openxlsx::write.xlsx(stat_test_df, "HASIL ESTIMASI ADF PP_TOTAL.xlsx") 



######## OPTIMAL LAG SELECTION ##########

diff_compdat <- comp_dat
diff_compdat <- diff_compdat[-1,]
# Calculate the first difference for each column except the time column
for (col in colnames(comp_dat)) {
  diff_compdat[[col]] <- diff(comp_dat[[col]])
}


lagselect <- VARselect(diff_compdat, lag.max = 7, type = "both")
lagselect$selection
lagselect_res <- data.frame("AIC" = lagselect$selection[1],"HQ" = lagselect$selection[2],
                            "SC" = lagselect$selection[3], "FPE" = lagselect$selection[4])

opt_lag <- 6


######## COINTEGRATION TEST ##########

ctest1t <- ca.jo(comp_dat_level, type = "trace", ecdet = "const", K = opt_lag-1)
summary(ctest1t) ## <- coint == 4


ctest1t@V ## Identify Long-rung relationship (Cointegrating Vectors)

ctest1t@W ## Identify Adjustment Coefficient

ctest1e <- ca.jo(comp_dat_level, type = "eigen", ecdet = "trend", K = opt_lag-1,spec = "longrun")
summary(ctest1e) ## <- coint == 2 (Decided to use eigen)

ctest1e@V
ctest1e@teststat

cajorls(ctest1e,r = 2)
summary(cajorls(ctest1e, r = 2)$rlm)

##PERSAMAAN JANGKA PENDEK
dat <- summary(a$rlm)
##PERSAMAAN JANGKA PANJANG
library(stargazer)
stargazer(a$beta, type = "html", out = "LongRun_total.html")



varest <- VAR(diff_compdat, p = opt_lag, type = "both")
plot(roots(varest))


################## VECM ####################

## VECM
options(digits=9)

Model1 <- VECM(diff_compdat, lag = opt_lag, r = 2, estim =("ML"))
summary(Model1)$coefficient

summary(Model1)
summary(rank.test(Model1))

res <- summary(Model1)


res$Pvalues
stargazer(summary(Model1)$coefMat, type = "html", out = "LongRun_total.html",star.char = c("+", "*", "**", "***"),
          star.cutoffs = c(0.1, 0.05, 0.01, 0.001),
          notes = c("+ p<0.1; * p<0.05; ** p<0.01; *** p<0.001"), 
          notes.append = T)
logLik(Model1)



t(Model1$model.specific$beta)

tes <- data.frame(summary(Model1)$coefMat)

Model1
coefs_all <- summary(Model1)$coefMat
# coefs_t <- summary(Model1)
# stargazer(coefs_t, type = "html", out = "VECM_total.html")
# options(scipen = 3)
coefs_all[grep("ECT", rownames(coefs_all)),]
ECT_total <- stargazer(coefs_all[grep("ECT", rownames(coefs_all)),], type = "html", out = "ECT_total.html")





varmod <- vec2var(ctest1t, r = 2)



vecm_beef <- irf(varmod, impulse = "beef", n.ahead = 20)
vecm_egg <- irf(varmod, impulse = "egg", n.ahead = 20)
vecm_cayenne <- irf(varmod, impulse = "cayenne", n.ahead = 20)
vecm_sugar <- irf(varmod, impulse = "sugar", n.ahead = 20)
vecm_shallot <- irf(varmod, impulse = "shallot", n.ahead = 20)
vecm_rice <- irf(varmod, impulse = "rice", n.ahead = 20)
vecm_chili <- irf(varmod, impulse = "red.chili", n.ahead = 20)
vecm_garlic <- irf(varmod, impulse = "garlic", n.ahead = 20)
vecm_oil <- irf(varmod, impulse = "cooking.oil", n.ahead = 20)
vecm_chicken <- irf(varmod, impulse = "chicken", n.ahead = 20)



Pattern1 <- list(mget(ls(pattern = "vecm_")))[[1]]
names(Pattern1) <- paste0("extr_", names(Pattern1))

irftest <- list()
for(i in names(Pattern1)){
  a <- list(extract_varirf(Pattern1[[i]]))
  irftest <- append(irftest,a)
}
names(irftest) <- names(Pattern1)



for(i in names(irftest)){
  for(j in 2:11){
    plotfile <- paste0("plot",colnames(irftest[[i]][j]),".png")
    test <- ggplot(irftest[[i]], aes(x=irftest[[i]][,1], y=irftest[[i]][,j], ymin=irftest[[i]][,j+10], ymax=irftest[[i]][,j+20])) +
      geom_hline(yintercept = 0, color="red") +
      geom_ribbon(fill="grey", alpha=0.2) +
      geom_line() +
      theme_light() +
      ggtitle(paste("Orthogonal impulse response from",sapply(strsplit(colnames(irftest[[i]])[j], "_"), function(x) x[2])))+
      ylab(paste0("\u0394(log(",sapply(strsplit(colnames(irftest[[i]])[j], "_"), function(x) tail(x, 1)),"))"))+
      xlab("Horizon") +
      theme(plot.title = element_text(size = 13, hjust=0.5),
            axis.title.y = element_text(size=11))
    ggsave(paste0("IRF Res/",plotfile), plot = test, width = 8, height = 6, dpi = 300)
  }
}



fevd(varmod, n.ahead = 50)