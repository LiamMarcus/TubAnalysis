# Load all librarys and load my custom functions
source(file="E:/Projects/Rutility/UtilityFunctions.R")

#Read all the tabs of the excel data file
dataFileName <- tclvalue(tkgetOpenFile())
myData<-importWorksheets(dataFileName)
names(myData)

#select the worksheet from the object myData
data<-myData$Gauge
names(data)

tbl_df(data)

#Clean-up of a gauge data table using tidyr
#this call works on a table of gauge data where the locations are column headings
#and the data are the gauge measurements.
data <- data %>%
    gather(location,gauge,B1:S9)

#Next, let's change all the gauges to mils using my fixGauge function, but letting
#dplyr's mutate verb apply the function

data <- data %>%
    mutate(gauge =fixGauge(gauge))

#this will put the table back into the original order using tidyr's spread()
#data <- data %>%
#    spread(location,gauge)

#now let's look at some trial summaries using
data %>%
    group_by(Profile,location) %>%
    summarize(aveGauge = mean(gauge))

data %>%
    group_by(Stack) %>%
    summarize(aveGauge = mean(gauge))


