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
library(seasonal)
library(forecast)
library(reshape)
library(RColorBrewer)




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
  a <- log(comp_dat2[i],base = 10)
  comp_dat <- append(comp_dat,a)
}
comp_dat <- data.frame(comp_dat)
comp_dat$Date <- comp_dat2$Date
comp_dat <- comp_dat[, c("Date", names(comp_dat)[names(comp_dat) != "Date"])]

######## COINTEGRATION ##########
comp_dat_level <- comp_dat
comp_dat_level$day <- weekdays(comp_dat_level$Date)
comp_dat_level <- comp_dat_level[comp_dat_level$day != "Saturday",]
comp_dat_level <- comp_dat_level[comp_dat_level$day != "Sunday",]
comp_dat_level <- comp_dat_level %>% dplyr::select(., c(-day))

#####################CHANGING DATA STRUCTURE
comp_dat2_untukdeseason <- comp_dat_level %>% dplyr::select(., c(-Date))
#### total = comp_dat_level %>% dplyr::select(., c(-Date))
#### pre = comp_dat_level[comp_dat_level$Date < "2022-09-03",] %>% dplyr::select(., c(-Date))
#### post = comp_dat_level[comp_dat_level$Date >= "2022-09-03",] %>% dplyr::select(., c(-Date))
comp_dat_level <- comp_dat_level %>% dplyr::select(., c(-Date))


comp_dat <- comp_dat_level    ######<-The estimation starts from log transformation data



######## UNIT ROOT TESTS ##########
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
  names(adf)[i] <- datname[i]
  names(adf_d1)[i] <- datname[i]
  names(pp)[i] <- datname[i]
  names(pp_d1)[i] <- datname[i]
}
stat_test <- list("adf" = melt(data.frame(adf)),"adf_d1" = melt(data.frame(adf_d1)),
                  "pp" = melt(data.frame(pp)),"pp_d1" = melt(data.frame(pp_d1)))

stat_test2_beforetreat <- stat_test %>% reduce(left_join, by = "variable")
colnames(stat_test2_beforetreat)[2:ncol(stat_test2_beforetreat)] <- names(stat_test)


######## DESEASONALIZE ##########
for(i in 1:ncol(comp_dat2_untukdeseason)){
  dtset2[[i]]<-ts(comp_dat2_untukdeseason[,i],start = c(year = 2018, month = 1, day = 1),frequency = 366)
}


deseason <- list()
for(i in names(dtset2)){
  a <- stl(dtset2[[i]],s.window = "periodic")
  a <- list(dtset2[[i]] - a$time.series[,1])
  deseason <- append(deseason,a)
}
names(deseason) <- names(dtset2)

if(nrow(comp_dat) == 1565){
  deseason[[i]] <- deseason[[i]]
}else if(nrow(comp_dat) == 1220){
  for(i in names(deseason)){
    deseason[[i]] <- deseason[[i]][c(1:1220)]
  }
}else{
  for(i in names(deseason)){
    deseason[[i]] <- deseason[[i]][c(1221:length(deseason[[i]]))]
  }
}

  
# if(nrow(comp_dat) == 1220){
#   for(i in names(deseason)){
#     deseason[[i]] <- deseason[[i]][c(1:1220)]
#   }
# }else{
#   for(i in names(deseason)){
#     deseason[[i]] <- deseason[[i]][c(1221:length(deseason[[i]]))]
#   }
# }




######## UNIT ROOT (AFTER DESEASONALIZED) ##########
dtset2 <- list()
adf <- list()
pp <- list()
adf_d1 <- list()
pp_d1 <- list()
for(i in 1:length(deseason)){
  a <- adf.test(deseason[[i]])
  b <- adf.test(diff(deseason[[i]]))
  
  c <- pp.test(deseason[[i]])
  d <- pp.test(diff(deseason[[i]]))
  
  adf <- append(adf,a["p.value"])
  adf_d1 <- append(adf_d1,b["p.value"])
  pp <- append(pp,c["p.value"])
  pp_d1 <- append(pp_d1,d["p.value"])
  names(deseason)[i] <- colnames(comp_dat)[i]
}

for(i in 1:length(adf)){
  names(adf)[i] <- datname[i]
  names(adf_d1)[i] <- datname[i]
  names(pp)[i] <- datname[i]
  names(pp_d1)[i] <- datname[i]
}
stat_test <- list("adf" = melt(data.frame(adf)),"adf_d1" = melt(data.frame(adf_d1)),
                  "pp" = melt(data.frame(pp)),"pp_d1" = melt(data.frame(pp_d1)))

stat_test2_afterseason <- stat_test %>% reduce(left_join, by = "variable")
colnames(stat_test2_afterseason)[2:ncol(stat_test2_afterseason)] <- names(stat_test)


######## DETRENDING ##########
detrend <- list()
for(i in names(deseason)){
  a <- list(resid(lm(coredata(deseason[[i]]~index(deseason[[i]])))))
  detrend <- append(detrend, a)
}
names(detrend) <- names(deseason)
detrend <- data.frame(detrend)

######## UNIT ROOT (AFTER DESEASONALIZE AND DETREND) ##########
dtset2 <- list()
adf <- list()
pp <- list()
adf_d1 <- list()
pp_d1 <- list()
for(i in 1:ncol(detrend)){
  dtset2[[i]]<-ts(detrend[,i],start = c(year = 2018, month = 1, day = 1),frequency = 366)
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
  names(adf)[i] <- datname[i]
  names(adf_d1)[i] <- datname[i]
  names(pp)[i] <- datname[i]
  names(pp_d1)[i] <- datname[i]
}
stat_test <- list("adf" = melt(data.frame(adf)),"adf_d1" = melt(data.frame(adf_d1)),
                  "pp" = melt(data.frame(pp)),"pp_d1" = melt(data.frame(pp_d1)))

stat_test2_aftertreat <- stat_test %>% reduce(left_join, by = "variable")
colnames(stat_test2_aftertreat)[2:ncol(stat_test2_aftertreat)] <- names(stat_test)

######## EXPORT stat_test2 ##########
stat_test_excel <- list("stat_test2_afterseason" = stat_test2_afterseason, "stat_test2_afterdetrend" = stat_test2_aftertreat, "stat_test2_beforetreat" = stat_test2_beforetreat)

if(nrow(comp_dat) == 1565){
  openxlsx::write.xlsx(stat_test_excel, "HASIL ESTIMASI ADF PP_TOTAL.xlsx", overwrite = T) 
}else if(nrow(comp_dat_level) == 1220){
  openxlsx::write.xlsx(stat_test_excel, "HASIL ESTIMASI ADF PP_TOTAL_PRE.xlsx", overwrite = T) 
}else{
  openxlsx::write.xlsx(stat_test_excel, "HASIL ESTIMASI ADF PP_TOTAL_POST.xlsx", overwrite = T) 
}


######## ZA_TEST ##########

perform_za_test <- function(ts_data, ts_name) {
  za_test <- ur.za(ts_data, model = "both", lag = 14)
  teststat <- za_test@teststat
  cval <- za_test@cval[2]
  bpoint <- za_test@bpoint
  dat <- data.frame("teststat" = teststat, "cval" = cval, "bpoint" = bpoint)
  return(dat)
}

zatest_deseason <- data.frame()
for (ts_name in names(deseason)) {
  a <- perform_za_test(deseason[[ts_name]], ts_name)
  zatest_deseason <- rbind(zatest_deseason,a)
}
zatest_deseason$final <- ifelse(zatest_deseason$teststat > zatest_deseason$cval, "NON stat","stat")
zatest_deseason$name <- datname ### <-HASIL STATTEST

zatest_detrend <- data.frame()
for (ts_name in names(detrend)) {
  a <- perform_za_test(detrend[[ts_name]], ts_name)
  zatest_detrend <- rbind(zatest_detrend,a)
}
zatest_detrend$final <- ifelse(zatest_detrend$teststat > zatest_detrend$cval, "NON stat","stat")
zatest_detrend$name <- datname ### <-HASIL STATTEST

zatest_baselog <- data.frame()
for (ts_name in names(comp_dat)) {
  a <- perform_za_test(comp_dat[[ts_name]], ts_name)
  zatest_baselog <- rbind(zatest_baselog,a)
}
zatest_baselog$final <- ifelse(zatest_baselog$teststat > zatest_baselog$cval, "NON stat","stat")
zatest_baselog$name <- datname ### <-HASIL STATTEST


######## EXPORT ZATEST ##########
zatest_full <- list("zatest_deseason" = zatest_deseason, "zatest_baselog" = zatest_baselog, "zatest_detrend" = zatest_detrend)

if(nrow(comp_dat) == 1565){
  openxlsx::write.xlsx(zatest_full, "zatest_full.xlsx", overwrite = T)
}else if(nrow(comp_dat_level) == 1220){
  openxlsx::write.xlsx(zatest_full, "zatest_full_PRE.xlsx", overwrite = T)
}else{
  openxlsx::write.xlsx(zatest_full, "zatest_full_POST.xlsx", overwrite = T)
}


# ######## FIRST DIFF (BASE) ##########
# diff_compdat <- comp_dat
# diff_compdat <- diff_compdat[-1,]
# # Calculate the first difference for each column except the time column
# for (col in colnames(comp_dat)) {
#   diff_compdat[[col]] <- diff(comp_dat[[col]])
# }

######## FIRST DIFF (DESEASON) ##########
diff_compdat <- data.frame(deseason)
diff_compdat <- diff_compdat[-1,]
# Calculate the first difference for each column except the time column
for (col in colnames(data.frame(deseason))) {
  diff_compdat[[col]] <- diff(data.frame(deseason)[[col]])
}
# 
# ######## FIRST DIFF (DETREND) ##########
# diff_compdat <- detrend
# diff_compdat <- diff_compdat[-1,]
# # Calculate the first difference for each column except the time column
# for (col in colnames(detrend)) {
#   diff_compdat[[col]] <- diff(detrend[[col]])
# }

######## LAG SELECTION ##########
fix_dat <- diff_compdat



lagselect <- vars::VARselect(fix_dat, lag.max = 14, type = "both")
lagselect
opt_lag <- 6

######## VAR EST & DIAGNOSTIC TESTS ##########
varest <- VAR(fix_dat, p = opt_lag, type = "both")
roots(varest)

fevd_est <- fevd(varest,n.ahead = 60)


ser11 <- serial.test(varest, lags.pt = opt_lag, type = "PT.asymptotic")
ser11$serial
norm1 <-normality.test(varest)
norm1$jb.mul
arch1 <- arch.test(varest, lags.multi = opt_lag)
arch1$arch.mul


for(i in 1:length(fevd_est)){
  fevd_est[[i]] <- data.frame(fevd_est[[i]])
  nama[i] <- paste0("fevd_",colnames(comp_dat)[i])
  names(fevd_est)[[i]] <- nama[i]
}

fevd_fix <- list()
for(i in 1:length(fevd_est)){
  fevd_fix[[i]] <- data.frame(fevd_est[[i]] %>% slice(c(1:30))) %>% mutate(period = c(1:30))
  names(fevd_fix)[i] <- names(fevd_est)[i]
}
fevd_graph <- list()
for(i in names(fevd_fix)){
  a <- fevd_fix[[i]] %>% pivot_longer(cols = -period, names_to = "commodity", values_to = "value") %>% list()
  fevd_graph <- append(fevd_graph,a)
}
names(fevd_graph) <- stringr::str_remove(names(fevd_fix),pattern = "fevd_")


for(i in 1:length(fevd_graph)){
  plotfile_FEVD <- paste0("FEVD_",names(fevd_graph)[[i]],".png")
  plot_fevd <- ggplot(fevd_graph[[i]], aes(fill=commodity, y=value, x=period)) + 
    ggtitle(paste0("Variance Decomposition of ",names(fevd_graph)[[i]])) +
    ylab("FEVD") + xlab("horizon (days)") + 
    theme(plot.title = element_text(size = 16,face="bold"), axis.title.y = element_text(size=12), 
          axis.title.x = element_text(size = 12),
          panel.background = element_rect(fill='transparent'),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          legend.background = element_rect(fill='transparent'),
          legend.title = element_text(size = 12,face = "bold"),
          legend.box.background = element_rect(fill='transparent')) +
    geom_bar(position="fill", stat="identity",width = 0.90) + coord_fixed(30)
  if(nrow(comp_dat) == 1565){
    ggsave(paste0("FEVD/TOTAL/",plotfile_FEVD), plot = plot_fevd, width = 8, height = 6, dpi = 300)
  }else if(nrow(comp_dat_level) == 1220){
    ggsave(paste0("FEVD/PRE/",plotfile_FEVD), plot = plot_fevd, width = 8, height = 6, dpi = 300)
  }else{
    ggsave(paste0("FEVD/POST/",plotfile_FEVD), plot = plot_fevd, width = 8, height = 6, dpi = 300)
  }
}




if(nrow(comp_dat) == 1565){
  openxlsx::write.xlsx(fevd_fix,"fevd_est.xlsx",overwrite = T)
}else if(nrow(comp_dat_level) == 1220){
  openxlsx::write.xlsx(fevd_fix,"fevd_est PRE.xlsx",overwrite = T)
}else{
  openxlsx::write.xlsx(fevd_fix,"fevd_est POST.xlsx",overwrite = T)
}
       

irfbulk <- list()
for (i in names(fix_dat)){
  nama <- i
  a <- list(irf(varest, impulse = i, n.ahead = 30))
  irfbulk <- append(irfbulk,a)
} 

for (i in 1:length(irfbulk)){
  names(irfbulk)[i] <- paste0("irf_",names(fix_dat)[i])
  list2env(irfbulk[i],.GlobalEnv)
}


irftest <- list()
for(i in names(irfbulk)){
  a <- list(extract_varirf(irfbulk[[i]]))
  irftest <- append(irftest,a)
}
names(irftest) <- names(irfbulk)

irffix <- list()
for(i in 1:length(irftest)){
  irffix[[i]] <- data.frame(irftest[[i]][c(1:11)])
}


irf_test2 <- list()
for(i in 1:length(irffix)){
  a <- irffix[[i]] %>%
    pivot_longer(cols = -period, names_to = "commodity", values_to = "value") %>% mutate(commodity = gsub(".*_(.*)", "\\1", commodity)) %>% list() 
  irf_test2 <- append(irf_test2,a)
}
names(irf_test2) <- names(fevd_graph)

# Plot multiple lines using ggplot2
for(i in 1:length(irf_test2)){
  plotfile_IRF <- paste0("IRF",names(irf_test2)[[i]],".png")
  plot_IRF <- ggplot(irf_test2[[i]], aes(x = period, y = value, color =commodity)) +
    geom_line(size=0.7) +
    geom_hline(yintercept = 0) +
    theme(plot.title = element_text(size = 16,face="bold"), axis.title.y = element_text(size=12), 
          axis.title.x = element_text(size = 12),
          panel.background = element_rect(fill='transparent'),
          panel.grid.major = element_blank(),
          panel.grid.minor = element_blank(),
          legend.background = element_rect(fill='transparent'),
          legend.title = element_text(size = 16,face = "bold"),
          legend.box.background = element_rect(fill='transparent')) +
    scale_color_brewer(palette = "Paired") +
    labs(title = paste0("Orthogonal impulse response from ",names(irf_test2)[[i]]), x = "horizon (days)", y = "response of 1 s.d shock")
  if(nrow(comp_dat) == 1565){
    ggsave(paste0("IRF Res_Diff Deseasonal/multiple lines plot/TOTAL/",plotfile_IRF), plot = plot_IRF, width = 10, height = 10, dpi = 300)
  }else if(nrow(comp_dat_level) == 1220){
    ggsave(paste0("IRF Res_Diff Deseasonal/multiple lines plot/PRE/",plotfile_IRF), plot = plot_IRF, width = 10, height = 10, dpi = 300)
  }else{
    ggsave(paste0("IRF Res_Diff Deseasonal/multiple lines plot/POST/",plotfile_IRF), plot = plot_IRF, width = 10, height = 10, dpi = 300)
  }
}




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
    if(nrow(comp_dat) == 1565){
      ggsave(paste0("IRF Res_Diff Deseasonal/TOTAL/",plotfile), plot = test, width = 8, height = 6, dpi = 300)
    }else if(nrow(comp_dat_level) == 1220){
      ggsave(paste0("IRF Res_Diff Deseasonal/PRE/",plotfile), plot = test, width = 8, height = 6, dpi = 300)
    }else{
      ggsave(paste0("IRF Res_Diff Deseasonal/POST/",plotfile), plot = test, width = 8, height = 6, dpi = 300)
    }
    
  }
}

