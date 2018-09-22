#Input the Expense Sheet
expsheet<-tidyxl::xlsx_cells("Expenses 2017.xlsx")

#Write code for any subsetting that needs to be done
#----------------Write Here----------------------
#
#
#-----------------Carry on-----------------------

#For each month, remove different sections...
#Start by getting names of the months in the DF
mth<-levels(as.factor(expsheet$sheet))

#Create a subset of transactions with values of month, row, column, numeric and comment
trans<-expsheet
#browser()
for (i in 1:length(mth)) {
    #4. Remove values for current month where row>the number in which Total is present for col1
    #Selecting row value for which the above condition is met for present month
    vrow<-trans$row[which(trans$sheet==mth[i] & trans$col==1 & trans$character=="Total")]
    rrow<-which(trans$sheet==mth[i] & trans$row>=vrow)
    #browser()
    trans <- trans[-rrow,]
    #browser()
    #5. Remove all obs with col >= the col for row2 and first 3 letters are SUM
    vcol<-trans$col[which(trans$sheet==mth[i] & trans$row==2 & substr(trans$formula,1,3)=="SUM")]
    rcol<-which(trans$sheet==mth[i] & trans$col>=vcol)
    #browser()
    trans <- trans[-rcol,]
    #browser()
}

#1. remove all obs where row=1
trans<-trans[-which(trans$row==1),]
#browser()
#2. remove all obs from col 1
trans<-trans[-which(trans$col==1),]
#browser()
#3. delete blank entries
trans <- trans[-which(is.na(trans$numeric)),names(trans) %in% c("sheet","row","col","numeric","comment","date","character")]
#browser()

#now time to get the date and category values from the bigger sheet
for (j in 1:nrow(trans)) {
    trans$date[j] <- expsheet$date[which(expsheet$sheet==trans$sheet[j] & expsheet$row==1 & expsheet$col==trans$col[j])]
    trans$character[j] <- expsheet$character[which(expsheet$sheet==trans$sheet[j] & expsheet$col==1 & expsheet$row==trans$row[j])]
}


#Creating a DF with the Date, Category and Amount spent in that category that day
#Correction. Changed to only by date
catDF <- aggregate(trans$numeric, by=list(date=trans$date),FUN="sum")


#Calculating Pro-rata projected amounts for each month and category.
#proDF <- data.frame(month=character(0),category=character(0),pltotal=numeric(0),projdaily=numeric(0))
proDF <- expsheet
for (k in 1:length(mth)) {
    #begin stripping values not revelant
    #browser()
    vrow<-proDF$row[which(proDF$sheet==mth[k] & proDF$col==1 & proDF$character=="Total")]
    proDF<-proDF[-which(proDF$sheet==mth[k] & proDF$row>=vrow),]
    proDF<-proDF[-which(proDF$sheet==mth[k] & proDF$col>1),]
    proDF<-proDF[-which(proDF$sheet==mth[k] & proDF$row<2),]
}
proDF <- proDF[,names(proDF) %in% c("sheet","row","col","numeric","character","height")]
dinmonth <- c(31,28,31,30,31,30,31,31,30,31,30,31)
#pick the planned total values for the remaining observations

for (k in 1:nrow(proDF)) {
    colnum <- expsheet$col[which(expsheet$sheet==proDF$sheet[k] & expsheet$row==1 & substr(expsheet$character,1,4)=="Plan")]
    proDF$numeric[k] <- expsheet$numeric[which(expsheet$sheet==proDF$sheet[k] & expsheet$row==proDF$row[k] & expsheet$col==colnum)]
    proDF$date[k] <- expsheet$date[which(expsheet$sheet==proDF$sheet[k] & expsheet$row==1 & expsheet$col==2)]
    #proDF$height <- round(proDF$numeric/dinmonth[which(mth==proDF$sheet[k])],0)
}

colnames(proDF) <- c("month","row","column","total_proj","category","daily_proj")
proDF <- proDF[,c(1,2,3,5,4,6)]
#Just realized after writing this code that its useless because the category daily projections are a very bad
#indicator of deviation. Might use this in another analysis in the near future. Fow now, will
#just use the aggregate function to sum up totals for each month on a daily basis.

dailyprojDF <- aggregate(proDF$daily_proj, by=list(month=proDF$month), FUN="sum")

#Will now start by calculating the difference in day's expense and the projected expense for that month.
#Step one: Calculate the summed daily expenditure (month, date, expense)

dailyExp <- aggregate(numeric ~ sheet + date, data=trans, FUN="sum")

#Now add the values from prorata to this one
dailyExp <- merge(dailyExp,dailyprojDF, by.x = c("sheet"), by.y = c("month"))
dailyExp$devExp <- dailyExp$numeric - dailyExp$x
dailyExp$pDev <- dailyExp$devExp/dailyExp$x

#Now order the table by decreasing order of deviation
dailyExp <- dailyExp[order(-dailyExp$pDev),]
dailyExp$cumil <- numeric(nrow(dailyExp))

#Now adding the cumilative of excess expenditure
dailyExp$cumil[1]=dailyExp$devExp[1]
for (i in 2:nrow(dailyExp)) {
    dailyExp$cumil[i]=dailyExp$devExp[i]+dailyExp$cumil[i-1]
}

#Going to look at the following plots for an indicator of how deviation is
#spread through the year
plot(dailyExp$date,dailyExp$pDev)
plot(dailyExp$date,dailyExp$devExp)

#Now attempting to do a root cause analysis. I want to assign ONE expense 
#category to each date value. I am assmung that for each day there must have been
#one category expense worth pointing out. For now, let me take that as the one 
#who's absolute value is the highest

dailyExp$maxcat <- character(nrow(dailyExp))

#Run a for loop for each fow of the dailyExp

for (i in 1:nrow(dailyExp)) {
    dailyExp$maxcat[i]=subset(trans[order(trans$numeric, decreasing = TRUE),], date == dailyExp$date[i])$character[1]
}

#Now try to figure out the root cause of increased expenses
library(ggplot2)
ggplot(dailyExp[abs(dailyExp$pDev)<10,], aes(x=date, y=pDev)) + geom_point(aes(color=maxcat))

ggplot(dailyExp[abs(dailyExp$pDev)>3 & abs(dailyExp$pDev)<15 & dailyExp$maxcat %in% c("Misc","Socializing", "Grooming and Health"),], aes(x=date, y=devExp)) + geom_point(aes(color=maxcat, size=pDev)) + labs(title="Category-wise Expense Contributors", x="Month of 2017", y="Rupees over Budget", color="Expense Category", size="Times Budgeted")

#now looking at the graph of expenses accross months and dates
ggplot(data=subset(trans, numeric>0 & numeric<=5000), aes(x=jitter(as.numeric(format(date,"%m"))),y=jitter(as.numeric(format(date,"%d"))), color=character, size=numeric)) + geom_point(alpha=0.5)

#previous version without jitter
ggplot(data=subset(trans, numeric>0 & numeric<=5000), aes(x=format(date,"%m"),y=format(date,"%d"), color=character, size=numeric)) + geom_point(alpha=0.5)

counter <- character()
test2 <- character()
for (i in 1:length(test)) {
    for (j in 2:length(test[[i]])) {
        test2 <- rbind(test2,as.character(test[[i]][[j]]))
        counter <- rbind(counter,as.character(i))
        #browser()
    }
}
test2 <- gsub("\r","",test2)

test3 <- matrix(nrow = length(counter),ncol=5)
test2 <- strsplit(test2,"=")
for (i in 1:length(counter)) {
    test3[i,1] <- counter[i]
    test3[i,2] <- test2[[i]][[1]]
    test3[i,3] <- test2[[i]][[2]]
}

for (i in 1:nrow(test3)) {
    test3[i,4] <- as.character(trans$character[as.numeric(test3[i,1])])
    test3[i,5] <- trans$date[as.numeric(test3[i,1])]
}

test3 <- as.data.frame(test3, stringsAsFactors=FALSE)
test3$V5 <- as_datetime(as.numeric(test3$V5))
test3 <- test3[,c(5,3,4,2)]
colnames(test3) <- c("Date","Amount","Category","Comment")



prorat <- function(ipdate){
    ipmonth <- as.numeric(format(ipdate,"%m"))
    ipyear <- as.numeric(format(ipdate,"%y"))
    if(ipmonth == 2 & ipyear%%4 == 0){
        return(29)
    } else if(ipmonth == 2 & ipyear%%4 != 0){
        return(28)
    } else {
        daycount <- c(31,28,31,30,31,30,31,31,30,31,30,31)
        return(daycount[ipmonth])
    }
}
test3$pMthBdg <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    test3$pMthBdg[i] <- test3$Amount[i]/sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)
}

test3$pMthCatBdg <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    test3$pMthCatBdg[i] <- test3$Amount[i]/sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)
}

test3$pDepCatMthBfr <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- sum(subset(test3, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE)
    denomenator <- sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)
    test3$pDepCatMthBfr[i]<- numerator/denomenator
}

test3$pDepCatMthAft <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- sum(subset(test3, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE) + test3$Amount[i]
    denomenator <- sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)
    test3$pDepCatMthAft[i]<- numerator/denomenator
}

# Feature no. 4a %age depletion of Monthly budget before this expense (pDepMthBfr)
test3$pDepMthBfr <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- sum(subset(test3, format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE)
    denomenator <- sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)
    test3$pDepMthBfr[i]<- numerator/denomenator
}

# Feature no. 4b %age depletion of Monthly budget after this expense (pDepMthAft)
test3$pDepMthAft <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- sum(subset(test3, format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE) + test3$Amount[i]
    denomenator <- sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)
    test3$pDepMthAft[i]<- numerator/denomenator
}

test3$pProCat <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- test3$Amount[i]
    denomenator <- sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i])
    test3$pProCat[i]<- numerator/denomenator
}


# Feature no. 5b %age of pro-rata daily amount(pProDly)
test3$pProDly <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- test3$Amount[i]
    denomenator <- sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i])
    test3$pProDly[i]<- numerator/denomenator
}


test3$pDepCumProBfr <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- sum(subset(test3, format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE)
    denomenator <- (sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))*(as.numeric(format(test3$Date[i],"%d")))
    test3$pDepCumProBfr[i]<- numerator/denomenator
}


test3$pDepCumProAft <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- sum(subset(test3, format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE) + test3$Amount[i]
    denomenator <- (sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))*(as.numeric(format(test3$Date[i],"%d")))
    test3$pDepCumProAft[i]<- numerator/denomenator
}

# Feature no. 6c %age depletion of the cumilative pro-rata category budget before the expense(pDepCumCatProBfr)
test3$pDepCumCatProBfr <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- sum(subset(test3, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE)
    denomenator <- (sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))*(as.numeric(format(test3$Date[i],"%d")))
    test3$pDepCumCatProBfr[i]<- numerator/denomenator
}

# Feature no. 6d %age depletion of the cumilative pro-rata category budget after the expense(pDepCumCatProAft)
test3$pDepCumCatProAft <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- sum(subset(test3, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE) + test3$Amount[i]
    denomenator <- (sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))*(as.numeric(format(test3$Date[i],"%d")))
    test3$pDepCumCatProAft[i]<- numerator/denomenator
}


test3$aDifCatProBfr <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    test3$aDifCatProBfr[i]<- sum(subset(test3, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE) - (sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))*(as.numeric(format(test3$Date[i],"%d")))
}


# Feature no. 7b Absolute Difference from overall pro-rata before expense(aDifProBfr)
test3$aDifProBfr <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    test3$aDifProBfr[i]<- sum(subset(test3, format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE) - (sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))*(as.numeric(format(test3$Date[i],"%d")))
}

# Feature no. 7a Absolute cumilative Difference from category pro-rata before expense(aDifCumCatProBfr)
test3$aDifCumCatProBfr <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    test3$aDifCumCatProBfr[i]<- sum(subset(test3, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE) - (sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))*(as.numeric(format(test3$Date[i],"%d")))
}

# Feature no. 7b Absolute cumilative Difference from overall pro-rata before expense(aDifCumProBfr)
test3$aDifCumProBfr <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    test3$aDifCumProBfr[i]<- sum(subset(test3, format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE) - (sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))*(as.numeric(format(test3$Date[i],"%d")))
}

# Feature no. 7c Absolute Difference from category pro-rata(aDifCatPro)
test3$aDifCatPro <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    test3$aDifCatPro[i]<- test3$Amount[i] - (sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))
}

# Feature no. 7d Absolute Difference from overall pro-rata(aDifPro)
test3$aDifPro <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    test3$aDifPro[i]<- test3$Amount[i] - (sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))
}