extractMonths <- function(inputDF){
    totalMonths <- length(inputDF$data)
    finalDF <- NULL
    for(i in 1:totalMonths){
        #browser()
        finalDF <- rbind(finalDF,cbind(inputDF$data[[i]],mname=names(inputDF$data)[i]))
        #browser()
    }
    write.csv(finalDF,file = "result.csv")
    return(finalDF)
}

giveExp <- function(inputDF,splitvar){
    listMonth <- split.data.frame(inputDF,inputDF[[splitvar]])
    for(i in 1:length(listMonth)){
        tempDF <- NULL
        tempDF <- listMonth[[i]]
        tempDF <- tempDF[tempDF$row != 1,]
        tempDF <- tempDF[tempDF$row<tempDF$row[which(tempDF$character=="Total" & tempDF$col==1)],]
        colsort <- sort(unique(tempDF$col))
        tempDF <- tempDF[! tempDF$col %in% c(colsort[1],colsort[length(colsort)],colsort[length(colsort)-1]),]
        tempDF <- tempDF[tempDF$data_type != "blank",]
        listMonth[[i]] <- tempDF
    }
    listMonth <- do.call("rbind",listMonth)
    listMonth <- listMonth[,c("address","row","col","content","formula","numeric","date","comment","mname")]
    return(listMonth)
}

assignDate <- function(dfsmall,dfbig){
    for(i in 1:nrow(dfsmall)){
        #browser()
        dfsmall$date[i]<-dfbig$date[dfbig$row==1 & dfbig$col==dfsmall$col[i] & dfbig$mname==dfsmall$mname[i]]
        dfsmall$category[i]<-dfbig$character[dfbig$col==1 & dfbig$row==dfsmall$row[i] & dfbig$mname==dfsmall$mname[i]]
    }
    return(dfsmall)
}