#Working Directory
setwd("C://Users//gcarlstedt//OneDrive - Tenaska, Inc//Documents//ERCOT//Requests//Orgination//LRS_ERS")
#Path where files are stored
path <- "C://Users//gcarlstedt//OneDrive - Tenaska, Inc//Documents//ERCOT//Requests//Orgination//LRS_ERS"
#Get all revelvant files within folder
files <- list.files(path=path, pattern="^RTM_",full.names = FALSE)
#---------------------------------------------------------------------------------------------------------------------------------------------------
#Added to ensure only raw .xlsx RTM files are read (As uploaded to ERCOT website)
xlsx <- files[grep(".xlsx", files)]
#---------------------------------------------------------------------------------------------------------------------------------------------------
#Read files to data frame
#---> Need to read all worksheets within workbooks as well, then compile
data <- foreach(i = 1:length(xlsx)) %do% {
       
         sheets <- excel_sheets(xlsx[i]) #Reading all sheets withing excel workbook
         table <- map_df(sheets, ~read_excel(xlsx[i], sheet = .x, skip = 8)) #Since consistent formatting across sheets, can easily map all sheets to data frame
         
         }
data_bind <- data.table::rbindlist(data, use.names=TRUE, fill = TRUE) # Map each data frame to single data table
data_subset <- data_bind[data_bind$RTDNCLR > 0 | data_bind$RTDERS >0,]
#Checking for NA values in data table
sapply(data_subset, function(x) sum(is.na(x)))
#Convert NAs in columns to 0s; We know from later values for these columns values should be 0
data_subset[is.na(data_subset)] <- 0
#Function to get season from date
getSeason <- function(DATES) {
       
         #Using 2012 since it is a leap year
         WS <- as.Date("2012-12-15", format = "%Y-%m-%d") # Winter Solstice
         SE <- as.Date("2012-3-15",  format = "%Y-%m-%d") # Spring Equinox
         SS <- as.Date("2012-6-15",  format = "%Y-%m-%d") # Summer Solstice
         FE <- as.Date("2012-9-15",  format = "%Y-%m-%d") # Fall Equinox
         
           # Convert dates from any year to 2012 dates
           d <- as.Date(strftime(DATES, format="2012-%m-%d"))
           
             ifelse (d >= WS | d < SE, "Winter",
                                 ifelse (d >= SE & d < SS, "Spring",
                                                             ifelse (d >= SS & d < FE, "Summer", "Fall")))
           
          }
#Adding granularity to datetime
Year <- as.integer(year(data_subset$'SCED Timestamp'))
Month <- as.integer(month(data_subset$'SCED Timestamp'))
Day <- as.integer(day(data_subset$'SCED Timestamp'))
Season <- getSeason(data_subset[,data_subset$'SCED Timestamp'])
#Appending above columns to original data set and wriing to new data table
final_DT <- add_column(data_subset, Year = Year, Month = Month, Day = Day, Season = Season, .after = 2)
#Frequency of ERS and LRS deployment
#------------------------------------------------------------------------------------------------------------------------------------------------
#count_DT <- final_DT %>%
#  group_by(Year, Season) %>%
#  summarise(LRS = round(sum(RTDNCLR>0), 0),
#           ERS = ceiling(sum(RTDERS>0)))
#------------------------------------------------------------------------------------------------------------------------------------------------
count_LRS <- final_DT[final_DT$RTDNCLR >0,] %>%
     group_by(Year, Season) %>%
     summarise(LRS = n_distinct(Day))
count_ERS <- final_DT[final_DT$RTDERS >0,] %>%
     group_by(Year, Season) %>%
     summarise(ERS = n_distinct(Day))
#Convert to data tables
data.table::setDT(count_LRS)
data.table::setDT(count_ERS)
avg_DT <- final_DT %>%
     group_by(Year, Season) %>%
     summarise(LRS = mean(RTDNCLR),ERS = mean(RTDERS))
#Convert to data table
data.table::setDT(avg_DT) #Raw filtered data
#Getting all seasons from main data table with all data
full_seasons <- getSeason(data_bind[,data_bind$'SCED Timestamp'])
#Permutations for years and seasons
combinations <- data_bind %>% expand(Year, full_seasons)
#Grabbing needed data tables  
LRS_list <- list(combinations, count_LRS, avg_DT[,c(1,2,3)])
#Merging and reducing above data tables to desired output
LRS <- Reduce(function (combinations, count_LRS, avg_DT) 
       merge(combinations, count_LRS, avg_DT, by.x = c('Year','full_seasons'), by.y = c('Year','Season'), all.x = TRUE), LRS_list)
#Converting to data table
data.table::setDT(LRS)
#Renaming columns
data.table::setnames(LRS, 2:4, c('Season','Days Deployed','Avg. Deployed LRS (MW)'))
#Setting all 0 to NAs, 0's are just place holder (Help with graphing)
LRS[LRS == 0] <- NA
#Adding helper column
LRS$'Helper' <- paste(LRS$Season, LRS$Year)
#Forece level/ factor for helper column (When graphing, won't revert to alphabetical)
LRS$'Time Period' <- factor(LRS$'Helper', levels = LRS$Helper)
#Reorder and findal data table
LRS_cols <- colnames(LRS)
LRS <- LRS %>% select(LRS_cols[6],LRS_cols[3],LRS_cols[4])
#Grabbing needed data tables  
ERS_list <- list(combinations, count_ERS, avg_DT[,c(1,2,4)])
#Merging and reducing above data tables to desired output
ERS <- Reduce(function (combinations, count_ERS, avg_DT) 
       merge(combinations, count_ERS, avg_DT, by.x = c('Year','full_seasons'), by.y = c('Year','Season'), all.x = TRUE), ERS_list)
#Converting to data table
data.table::setDT(ERS)
#Renaming columns
data.table::setnames(ERS, 2:4, c('Season','Days Deployed','Avg. Deployed ERS (MW)'))
#Setting NA values to 0 instead as place holders 
ERS[ERS == 0] <- NA
#Adding helper column
ERS$'Helper' <- paste(ERS$Season, ERS$Year)
#Forece level/ factor for helper column (When graphing, won't revert to alphabetical)
ERS$'Time Period' <- factor(ERS$'Helper', levels = ERS$Helper)
#Reorder and findal data table
ERS_cols <- colnames(ERS)
ERS <- ERS %>% select(ERS_cols[6],ERS_cols[3],ERS_cols[4])
#------------------------------------------------------------------------------------------------------------------------------------------------
#Write data table to excel file 
write_xlsx(list("Raw Data" = final_DT,"LRS" = LRS, "ERS" = ERS),
                           "Historical Real-Time ORDC and Reliability Deployment Price Adders and Reserves_V2.xlsx") #Dont need to specify path because of setwd()
#Calling pdf command to start the plot
pdf(file="Deployment Data.pdf", width = 4, height = 4)
#ERS graph
top <- ggplot(LRS,aes(LRS$'Time Period',LRS$'Avg. Deployed LRS (MW)'))+
     geom_point(aes(size=LRS$`Days Deployed`*10),shape=21,fill="royalblue", stat="identity", alpha= 0.5)+
     geom_text(aes(label=LRS$`Days Deployed`),size=7)+
     scale_size_identity()+
     theme(panel.grid.major=element_line(linetype=2,color="black"),
                     axis.text.x=element_text(angle=90,hjust=1,vjust=0))+
     ggtitle("2016-2020 LRS Deployment (MW)")+
     ylab("Avg. Deployed LRS (MW)")+
     ylim(0,11.5)+
     xlab("Time Period")+
     theme(panel.background = element_blank())
#LRS graph
   bottom <- ggplot(ERS,aes(ERS$'Time Period',ERS$'Avg. Deployed ERS (MW)'))+
     geom_point(aes(size=ERS$`Days Deployed`*10),shape=21,fill="slategray", stat="identity", alpha= 0.5)+
     geom_text(aes(label=ERS$`Days Deployed`),size=5, check_overlap = TRUE, fontface = 'bold')+
     scale_size_identity()+
     theme(panel.grid.major=element_line(linetype=2,color="black"),
            +         axis.text.x=element_text(angle=90,hjust=1,vjust=0))+
     ggtitle("2016-2020 ERS Deployment (MW)")+
     ylab("Avg. Deployed ERS (MW)")+
     ylim(415,420)+
     xlab("Time Period")+
     theme(panel.background = element_blank())
figure <- ggarrange(top + rremove("xlab"), bottom,
                                           ncol = 1, nrow = 2, align = "v")
#Calling pdf command to start the plot
ggsave("Deployment Graphs_V2.pdf", width = 12, height = 8)

