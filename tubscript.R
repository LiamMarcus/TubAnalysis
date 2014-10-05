# Load all librarys and load my custom functions
source(file="E:/Projects/Rutility/UtilityFunctions.R")
library(qcc)

#Read all the tabs of the excel data file
dataFileName <- tclvalue(tkgetOpenFile())
myData<-importWorksheets(dataFileName)
names(myData)

# select the name of the excel file to save all tables.
saveFileName <- tclvalue(tkgetSaveFile())

#Summarize the fresh and aged data -- no aggregation
dataFresh<-myData$"Fresh"
tbl_df(dataFresh)

freshSummary <- dataFresh %>%
    select(Weight:RimCrush) %>%
    summarise_each(funs(mean))

dataAged<-myData$"7-day"
tbl_df(dataAged)

agedSummary <- dataAged %>%
    select(Weight:RimCrush) %>%
    summarise_each(funs(mean))

fullSummary <- rbind(freshSummary,agedSummary)
tabName="fullSummary"
exportToExcel(fullSummary,saveFileName,tabName)

#Construst xbar charts and calculate Cp index for all aged parameters
attach(dataAged)
Weight<- qcc.groups(Weight,Stack)
q<-qcc(Weight, type="xbar",title="Control Chart \n Cool Whip 8oz Undersized Weight (g)",digits=4)
process.capability(q, spec.limits=c(15.4,17.1),target=16.25)

TopCrushPeak<- qcc.groups(TopCrushPeak,Stack)
q<-qcc(TopCrushPeak, type="xbar",title="Control Chart \n Cool Whip 8oz Undersized Top Crush Peak (lbf)",digits=4)
process.capability(q, spec.limits=c(15,55),target=35)

RimCrush<- qcc.groups(RimCrush,Stack)
q<-qcc(RimCrush, type="xbar",title="Control Chart \n Cool Whip 8oz Undersized Rim Crush (lbf)",digits=3)
process.capability(q, spec.limits=c(0.75,2))

RimDiameter<- qcc.groups(RimDiameter,Stack)
q<-qcc(RimDiameter, type="xbar",title="Control Chart \n Cool Whip 8oz Undersized Rim Diameter Deviation (in x 1000)",digits=3)
process.capability(q, spec.limits=c(-20,-10), target=-15)
detach(dataAged)

# USE THIS FOR GAUGE LATER ####################
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


