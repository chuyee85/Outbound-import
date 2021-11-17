library(xlsx)
library(stringr)
library(dplyr)
library(magrittr)
library(tidyr)
library(lubridate)
library(plotly)
library(zoo)


##路徑設定
dir = "~"
setwd(dir)

##讀取檔案
file.list <- list.files(pattern="*.xls")

##資料讀取及預先處理(去除NA、不需要雜項)
DailyImmigPosAll <- lapply(file.list, function(DailyImmigPosAll){
  DailyImmigPosAll <- read.xlsx2(DailyImmigPosAll, sheetIndex=1, startRow = 6)
  colname_ch <- colnames(DailyImmigPosAll[1:15])
  DailyImmigPosAll[colname_ch] <- sapply(DailyImmigPosAll[colname_ch],as.numeric)
  DailyImmigPosAll <- DailyImmigPosAll %>% drop_na()
})

##過濾總結列 
DailyImmigPosAll_final_version <- do.call(rbind.data.frame, DailyImmigPosAll) %>% unique()

#年線入境
DailyImmigPosAll_final_version_year  <- DailyImmigPosAll_final_version %>% filter(., 入出境日期>2000     & 入出境日期<9999) 
DailyImmigPosAll_final_version_year$入出境日期 <- DailyImmigPosAll_final_version_year$入出境日期 %>% as.character()
DailyImmigPosAll_final_version_year$入出境日期 <- as.Date(ISOdate(DailyImmigPosAll_final_version_year$入出境日期, 12, 31)) %>% as.yearmon() %>% format("%Y")
fig_year<- plot_ly(DailyImmigPosAll_final_version_year, type = 'scatter', mode = 'lines+markers', autosize = T) %>% 
  add_trace(x = ~入出境日期, y = ~入境.合計, name = '年線入境') %>% 
  add_trace(x = ~入出境日期, y = ~出境.合計, name = '年線出境') %>%
  layout(
    title = list(text ='中外人士入出境統計表入境人數',
                 y = 0.99),
    xaxis = list(
      title = "入出境時間",
      tickformat = "%Y%m%d",
      date_breaks = '14 days'
    ),
    yaxis = list(
      title = '入出境人數',
      range = c(0, max(DailyImmigPosAll_final_version_year$入境.合計,DailyImmigPosAll_final_version_year$出境.合計))
    ),
    plot_bgcolor = "#e5ecf6"
  )

#月線
DailyImmigPosAll_final_version_month <- DailyImmigPosAll_final_version %>% filter(., 入出境日期>200001   & 入出境日期<999999)
DailyImmigPosAll_final_version_month$入出境日期 <- DailyImmigPosAll_final_version_month$入出境日期 %>% as.character() %>% as.yearmon(.,"%Y%m") %>% as.Date()
fig_month<- plot_ly(DailyImmigPosAll_final_version_month, type = 'scatter', mode = 'lines+markers', autosize = T) %>% 
  add_trace(x = ~入出境日期, y = ~入境.合計, name = '月線入境') %>%
  add_trace(x = ~入出境日期, y = ~出境.合計, name = '月線出境') %>% 
  layout(
    title = list(text ='中外人士入出境統計表入境人數',
                 y = 0.99),    
    xaxis = list(
      title = "入出境時間",
      tickformat = "%Y%m"
    ),
      yaxis = list(
      title = '入出境人數',
      range = c(0, max(DailyImmigPosAll_final_version_month$入境.合計,DailyImmigPosAll_final_version_month$出境.合計))
    ),
    plot_bgcolor = "#e5ecf6"
  )

#日線
DailyImmigPosAll_final_version_day   <- DailyImmigPosAll_final_version %>% filter(., 入出境日期>20000000 & 入出境日期<99999999)
DailyImmigPosAll_final_version_day$入出境日期 <- DailyImmigPosAll_final_version_day$入出境日期 %>% as.character() %>% as.Date(., format("%Y%m%d"))

fig_date<- plot_ly(DailyImmigPosAll_final_version_day, type = 'scatter', mode = 'lines', autosize = T) %>% 
  add_trace(x = ~入出境日期, y = ~入境.合計, name = '日線入境') %>% 
  add_trace(x = ~入出境日期, y = ~出境.合計, name = '日線出境') %>%
  add_trace(x = ~入出境日期, y = ~(rollmeanr(入境.合計, 7, fill = NA)), name = 'MA7_import', line = list(color = "blue")) %>%
  add_trace(x = ~入出境日期, y = ~(rollmeanr(出境.合計, 7, fill = NA)), name = 'MA7_Outbound', line = list(color = "red")) %>%
  layout(
    title = list(text ='中外人士入出境統計表入境人數',
                 y = 0.99),  
    xaxis = list( 
      title = "入出境時間",
      tickangle=45,
      tickformat = "%Y%m%d"
      ),
      yaxis = list(
      title = '入出境人數',
      range = c(0, max(DailyImmigPosAll_final_version_day$入境.合計,DailyImmigPosAll_final_version_day$出境.合計))
    ),
    plot_bgcolor = "#e5ecf6"
  )


#################加入週別
weekdate<- read.csv("V:/Rserver/移民署入出境人數/WEEKDATE.csv", colClasses = c("Date","character")) %>%
  filter(between(date,as.Date("2017-01-01"),Sys.Date())) 
colnames(weekdate) <- c("入出境日期","年週")

DailyImmigPosAll_final_version_day_week <- left_join(DailyImmigPosAll_final_version_day, weekdate, by="入出境日期") %>% group_by(年週) %>% summarise(入境.合計 = sum(入境.合計),
                                                                                                     出境.合計 = sum(出境.合計)) %>% as.data.frame()
fig_week<- plot_ly(DailyImmigPosAll_final_version_day_week, type = 'scatter', mode = 'lines', autosize = T) %>% 
  add_trace(x = ~年週, y = ~入境.合計, name = '週線入境') %>% 
  add_trace(x = ~年週, y = ~出境.合計, name = '週線出境') %>%
  layout(
    title = list(text ='中外人士入出境統計表入境人數',
                 y = 0.99),  
    xaxis = list( 
      title = "入出境時間",
      tickangle=90,
      type='category'
    ),
    yaxis = list(
      title = '入出境人數',
      range = c(0, max(DailyImmigPosAll_final_version_day_week$入境.合計,DailyImmigPosAll_final_version_day_week$出境.合計))
    ),
    plot_bgcolor = "#e5ecf6"
  )

##存檔
path = paste0("~")
write.xlsx(DailyImmigPosAll_final_version_year, path, row.names = FALSE, sheetName = "年資料")
write.xlsx(DailyImmigPosAll_final_version_month, path, row.names = FALSE, sheetName = "月資料", append=TRUE)
write.xlsx(DailyImmigPosAll_final_version_day, path, row.names = FALSE, sheetName = "日資料", append=TRUE)
write.xlsx(DailyImmigPosAll_final_version_day_week, path, row.names = FALSE, sheetName = "週資料", append=TRUE)
saveWidget(fig_year, "~", selfcontained = F, libdir = "外部程式")
saveWidget(fig_month, "~", selfcontained = F, libdir = "外部程式")
saveWidget(fig_date, "~", selfcontained = F, libdir = "外部程式")
saveWidget(fig_week, "~", selfcontained = F, libdir = "外部程式")




