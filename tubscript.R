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


#Summarize the gauge data
data<-myData$"Gauge"
tbl_df(data)

#Identify the factors and fix some of the units on variables
data$Profile <- factor(data$Profile)

#Clean-up of a gauge data table using tidyr
#this call works on a table of gauge data where the locations are column headings
#and the data are the gauge measurements.
data <- data %>%
    gather(location,gauge,B1:S9) %>%
    select(Profile:gauge)

#Next, let's change all the gauges to mils using my fixGauge function, but letting
#dplyr's mutate verb apply the function
data <- data %>%
    mutate(gauge =fixGauge(gauge))

#calulate the 1-sigma values for the plot bars
gaugeSummary <- summaryCI(data,"gauge",c("Profile","location"))

#Prepare the data for plotting
gaugeSummary$Upper <- gaugeSummary$mean / 2
gaugeSummary$Lower <- 0 - gaugeSummary$Upper
gaugeSummary$ciUpper <- gaugeSummary$ciUpper / 2
gaugeSummary$ciLower <- 0 - gaugeSummary$ciUpper

#make the plot and save the image
xlabel="Measurement Location"
ylabel="mils"
plotTitle="Cool Whip 8oz Tub Gauge Distribution"
p4<-ggplot(gaugeSummary) +
    geom_ribbon(aes(x=location,ymin=Lower,ymax=Upper, group=1), fill="red", alpha=.4) +
    geom_errorbar(aes(x=location,ymin=ciLower,ymax=ciUpper), colour="black", width=0.1) +
    facet_grid(Profile ~ .) +
    xlab(xlabel) +
    ylab(ylabel) +
    ggtitle(plotTitle) +
    theme(axis.text.x = element_text(angle = 90))

plotFile=paste("GaugeComparison.CIplot.png",    sep="")
ggsave(p4,width=6,height=4,file=plotFile)
p4
