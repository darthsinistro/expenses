#I want to first parse the data from an Excel Sheet into R

expsheet<-tidyxl::xlsx_cells("Expenses 2017.xlsx")

#Now, we start to clean the expsheet file and turn it into a format
#that we can understand. I will begin to create two tables with the following formats
#{Date, Amount, Category, Comment} and {Month & Year, Category, Burget}

#For the first part, I will systematically clip off observations which contain data
#for those rows which are below the Total Row, then remove the ones with column totals.

#I'm going to run this loop for each of the months in the Workbook. Running this operation on a copy
#because I will need to fetch the date and category data from the original dataframe later on.

mth<-levels(as.factor(expsheet$sheet))
trans<-expsheet

for (i in 1:length(mth)) {
    vrow<-trans$row[which(trans$sheet==mth[i] & trans$col==1 & trans$character=="Total")]
    rrow<-which(trans$sheet==mth[i] & trans$row>=vrow)
    trans <- trans[-rrow,]
    vcol<-trans$col[which(trans$sheet==mth[i] & trans$row==2 & substr(trans$formula,1,3)=="SUM")]
    rcol<-which(trans$sheet==mth[i] & trans$col>=vcol)
    trans <- trans[-rcol,]
}

#Now removing all entries containing data for the first row and first column. Also, keeping
#only specific columns and discarding the rest.

trans<-trans[-which(trans$row==1),]
trans<-trans[-which(trans$col==1),]
trans <- trans[-which(is.na(trans$numeric)),names(trans) %in% c("sheet","row","col","numeric","comment","date","character")]

#Now getting the date and category values stores in the Excel file, storing them in date
#and character columns respectively.

for (j in 1:nrow(trans)) {
    trans$date[j] <- expsheet$date[which(expsheet$sheet==trans$sheet[j] & expsheet$row==1 & expsheet$col==trans$col[j])]
    trans$character[j] <- expsheet$character[which(expsheet$sheet==trans$sheet[j] & expsheet$col==1 & expsheet$row==trans$row[j])]
}

#Now extracting the values from the comments.
test <- strsplit(trans$comment, "\n")
counter <- character()
test2 <- character()

#Have to extract the values from the list formed after strsplit
for (i in 1:length(test)) {
    for (j in 2:length(test[[i]])) {
        test2 <- rbind(test2,as.character(test[[i]][[j]]))
        counter <- rbind(counter,as.character(i))
    }
}

#Eliminating some '\r' values
test2 <- gsub("\r","",test2)

#Now starting to separate the value and comment separated by '='
test3 <- matrix(nrow = length(counter),ncol=5)
test2 <- strsplit(test2,"=")
for (i in 1:length(counter)) {
    test3[i,1] <- counter[i]
    test3[i,2] <- test2[[i]][[1]]
    test3[i,3] <- test2[[i]][[2]]
}

#Now taking out the date and category values from the trans data frame

for (i in 1:nrow(test3)) {
    test3[i,4] <- as.character(trans$character[as.numeric(test3[i,1])])
    test3[i,5] <- trans$date[as.numeric(test3[i,1])]
}

test3 <- as.data.frame(test3, stringsAsFactors=FALSE)
test3$V5 <- as_datetime(as.numeric(test3$V5))
test3 <- test3[,c(5,3,4,2)]
colnames(test3) <- c("Date","Amount","Category","Comment")
test3$Amount <- as.numeric(test3$Amount)
test3$Category <- as.factor(test3$Category)

final_expsheet <- test3

#WARNING============Cleaning up the Workspace====================
#rm(list=setdiff(ls(), "final_expsheet"))

#Now moving on to extracting the Budget data {Date, Category, Amount}

#!!!!!!!!!!!!!Oh boy. I see there are plenty of corrections to be made here with test3, final_expsheet and expsheet!!!!!!!!!!!!!!!!1

proDF <- expsheet
for (k in 1:length(mth)) {
    #begin stripping values not revelant
    #browser()
    vrow<-proDF$row[which(proDF$sheet==mth[k] & proDF$col==1 & proDF$character=="Total")]
    proDF<-proDF[-which(proDF$sheet==mth[k] & proDF$row>=vrow),]
    proDF<-proDF[-which(proDF$sheet==mth[k] & proDF$col>1),]
    proDF<-proDF[-which(proDF$sheet==mth[k] & proDF$row<2),]
}
#pick the planned total values for the remaining observations
for (k in 1:nrow(proDF)) {
    colnum <- expsheet$col[which(expsheet$sheet==proDF$sheet[k] & expsheet$row==1 & substr(expsheet$character,1,4)=="Plan")]
    proDF$numeric[k] <- expsheet$numeric[which(expsheet$sheet==proDF$sheet[k] & expsheet$row==proDF$row[k] & expsheet$col==colnum)]
    proDF$date[k] <- expsheet$date[which(expsheet$sheet==proDF$sheet[k] & expsheet$row==1 & expsheet$col==2)]
    #proDF$height <- round(proDF$numeric/dinmonth[which(mth==proDF$sheet[k])],0)
}
proDF <- proDF[,c("date","character","numeric")]
colnames(proDF) <- c("Date","Category","Mth_Budget")

#I also want to be able to calculate a pro-rata budget, so, will design a function for that.

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

#I am now going to start generating the feature set for each of these expenses
#including the day, month, weekday, among others listed on (IX - 28/5/18 2304)

# Feature no. 1 %age of monthly budget (pMthBdg)
test3$pMthBdg <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    test3$pMthBdg[i] <- test3$Amount[i]/sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)
}

# Feature no. 2 %age of monthly category budget (pMthCatBdg)
test3$pMthCatBdg <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    test3$pMthCatBdg[i] <- test3$Amount[i]/sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)
}

# Feature no. 3a %age depletion of Monthly category budget before this expense (pDepCatMthBfr)
test3$pDepCatMthBfr <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- sum(subset(test3, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE)
    denomenator <- sum(subset(proDF, Category==test3$Category[i] & format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)
    test3$pDepCatMthBfr[i]<- numerator/denomenator
}

# Feature no. 3b %age depletion of Monthly category budget after this expense (pDepCatMthAft)
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

# Feature no. 5a %age of pro-rata category amount(pProCat)
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

# Feature no. 6a %age depletion of the cumilative pro-rata overall budget before the expense(pDepCumProBfr)
test3$pDepCumProBfr <- c(1:nrow(test3))
for (i in 1:nrow(test3)) {
    numerator <- sum(subset(test3, format(Date, "%m")==format(test3$Date[i],"%m") & Date < test3$Date[i])$Amount, na.rm = TRUE)
    denomenator <- (sum(subset(proDF, format(Date, "%m")==format(test3$Date[i],"%m"))$Mth_Budget, na.rm = TRUE)/prorat(test3$Date[i]))*(as.numeric(format(test3$Date[i],"%d")))
    test3$pDepCumProBfr[i]<- numerator/denomenator
}

# Feature no. 6b %age depletion of the cumilative pro-rata overall budget after the expense(pDepCumProAft)
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
