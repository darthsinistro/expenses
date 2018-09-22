#Ask for the file name and the output DF name 
#so I can run multiple non-conflicting iterations
giveFlat <- function(inputFile){
    #read the inputFile using the tidyXL function
    exp <- tidyxl::tidy_xlsx(inputFile)
    #extract the data for all months into one DF
    combinedDF <- data.frame()
    for(i in 1:length(exp$data)){
        tempDF<-exp$data[[i]]
        #-------------------------------------------------------#
        #convert unstructured data into structured form
        stripedDF<-tempDF
        #1. delete all obervations of the 1st row
        stripedDF <- stripedDF[!stripedDF$row==1,]
        #2. delete all observations where row>= the row value of the...
        #...cell where col=1 and value=total
        trow<-stripedDF$row[which(stripedDF$col==1 & stripedDF$character=="Total")]
        stripedDF <- stripedDF[!stripedDF$row>=trow,]
        #3. delete all obervations of the 1st col
        stripedDF <- stripedDF[!stripedDF$col==1,]
        #4. delete all obersvations of cols where row=2 and contains SUM
        tcol <- stripedDF$col[which(stripedDF$row==2 & substr(stripedDF$formula,1,3)=="SUM")]
        stripedDF <- stripedDF[!stripedDF$col>=tcol,]
        #5. delete blank entries
        stripedDF <- stripedDF[!is.na(stripedDF$numeric),]
        #-------------------------------------------------------#
        #take out useful values - extract date and expense category
        for(j in 1:nrow(stripedDF)){
            stripedDF$date[j]<-tempDF$date[which(tempDF$row==1 & tempDF$col==stripedDF$col[j])]
            stripedDF$type[j]<-tempDF$character[which(tempDF$col==1 & tempDF$row==stripedDF$row[j])]
        }
        #keep only Date, Category, Amount, Comments
        vars<-c("date","numeric","type","comment")
        stripedDF<-stripedDF[,vars]
        #start joining to combinedDF
        mth<-as.data.frame(rep(names(exp$data)[i],nrow(stripedDF)))
        tempDF <- cbind(mth,stripedDF)
        colnames(tempDF)[1]<-"Month"
        combinedDF <- rbind(combinedDF,tempDF)
    }
    return(combinedDF)
}