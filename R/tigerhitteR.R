#' Complete the hollow dataset
#'
#' Take time series dataset and fields, then refill the missing date records and other fields.
#' @param inPath A path which is the location of uncompleted dataset which must be xlsx file
#' @param sheet A worksheet name of the dataset
#' @param dateCol Date column
#' @param outPath A path where the location of xlsx file of completed dataset should be
#' @param fixedVec  A row of column number which should be kept same values with the original
#' @param pChanged The column number which should be changed to different value into new record.
#' @param pChangedNum The value of a specific column which should be put into the new record.
#' @import openxlsx zoo Hmisc
#' @importFrom utils head install.packages
#' @details Real time series sales dataset could be not continuous in 'date' field. e.g., monthly sales data is continuous,
#'  but discrete in daily data.
#'
#'  This hollow dataset is not complete for time series analysis. Function dateRefill.fromFile
#'  is a transformation which tranforms uncomplete dataset into complete dataset.
#' @author Will Kuan
#' @examples # Please refer to the examples of function dateRefill.fromData
#' @export
dateRefill.fromsFile <-
  function(inPath, sheet, dateCol, outPath, fixedVec, pChanged, pChangedNum)
  {
    if(!requireNamespace("openxlsx",quietly = TRUE)){
      install.packages("openxlsx"); requireNamespace("openxlsx", quietly = TRUE)
      #stop("Please install package 'openxlsx'. ")
    }else{
      requireNamespace("openxlsx", quietly = TRUE)
    }

    if(!requireNamespace("zoo", quietly = TRUE)){
      install.packages("zoo"); requireNamespace("zoo", quietly = TRUE)
      #stop("Please install package 'zoo'. ")
    }else{
      requireNamespace("zoo", quietly = TRUE)
    }

    if(!requireNamespace("Hmisc",quietly = TRUE)){
      install.packages("Hmisc"); requireNamespace("Hmisc", quietly = TRUE)
      #stop("Please install package 'Hmisc'. ")
    }else{
      requireNamespace("Hmisc", quietly = TRUE)
    }

    data <- openxlsx::read.xlsx(inPath, sheet)

    data[,dateCol] <- zoo::as.Date(data[,dateCol], origin = "1899-12-30")
    colNameVector <- colnames(data)

    colnames(data)[dateCol] <- "Date"

    data$Date <- as.POSIXlt(data$Date) # transform to POSIXlt type

    year.list <- levels(factor(data$Date$year + 1900))

    ### sorting data
    inc.order <- order(data$Date, decreasing = FALSE)
    data <- data[inc.order,]

    ### building an empty data frame
    final.data <- data.frame(data[,1:length(colNameVector)])
    final.data[,] <- NA
    final.data$Date <- as.POSIXlt(final.data$Date)

    year <- substr(data[1,2],1,4)
    origin <- paste(year, "-01-01", sep = "")
    origin <- zoo::as.Date(origin)
    diff <- zoo::as.Date(data[1,2])-origin
    rm(year)

    daySum <- 0
    for(x in year.list)
    {
      daySum <- daySum + Hmisc::yearDays(zoo::as.Date(paste(x, "-01-01", sep = "")))
    }
    daySum <- as.numeric(daySum - diff)

    final.data[1:daySum,] <- NA # remove first few null days because data is not start on 1/1
    final.data[,2] <- seq(data[1,2], by = "1 days", length.out = daySum)

    #### setting rownames
    rownames(data) <- c(1:nrow(data))

    my.index <- match(data$Date, as.POSIXlt(final.data$Date))
    final.data[my.index,] <- data[,]

    for(i in daySum:1)
    {
      if(!is.na(final.data[i,1])){
        tag <- i
        rm(i)
        break;
      }else{
        next;
      }
    }

    final.data <- head(final.data,tag)  # cutting off last empty records

    for(j in 1:nrow(final.data)){
      if(is.na(final.data[j,1])){
        final.data[j,fixedVec] <- data[1,fixedVec]
        final.data[j,pChanged] <- pChangedNum
      }
    }

    final.data$Date <- zoo::as.Date(final.data$Date) # for correcting date time in excel

    colnames(final.data) <- colNameVector

    ### output data into excel file
    openxlsx::write.xlsx(final.data, file = outPath)

  }

#' Complete the hollow dataset
#'
#' Take time series dataset and fields, then refill the missing date records and other fields.
#' @param data The data.frame dataset which is ready to be processed
#' @param dateCol Date column
#' @param fixedVec  A row of column number which should be kept same values with the original
#' @param pChanged The column number which should be changed to different value into new record.
#' @param pChangedNum The value of a specific column which should be put into the new record.
#' @import zoo Hmisc
#' @importFrom utils head install.packages
#' @return The dataset which is completed.
#' @details Real time series sales dataset could be not continuous in 'date' field. e.g., monthly sales data is continuous,
#'  but discrete in daily data.
#'
#'  This hollow dataset is not complete for time series analysis. Function dateRefill.fromFile
#'  is a transformation which tranforms uncomplete dataset into complete dataset.
#' @author Will Kuan
#' @examples # mydata <- data.example
#' # mydata.final <- dateRefill.fromData(data = mydata,dateCol = 2,fixedVec = c(3:10),
#' #                                     pChanged = 11,pChangedNum = 0)
#' @export
dateRefill.fromData <-
  function(data, dateCol, fixedVec, pChanged, pChangedNum)
  {
    if(!requireNamespace("zoo", quietly = TRUE)){
      install.packages("zoo"); requireNamespace("zoo", quietly = TRUE)
      #stop("Please install package 'zoo'. ")
    }else{
      requireNamespace("zoo", quietly = TRUE)
    }

    if(!requireNamespace("Hmisc",quietly = TRUE)){
      install.packages("Hmisc"); requireNamespace("Hmisc", quietly = TRUE)
      #stop("Please install package 'Hmisc'. ")
    }else{
      requireNamespace("Hmisc", quietly = TRUE)
    }

    data <- data.frame(data)

    data[,dateCol] <- zoo::as.Date(data[,dateCol], origin = "1899-12-30")
    colNameVector <- colnames(data)

    colnames(data)[dateCol] <- "Date"

    data$Date <- as.POSIXlt(data$Date) # transform to POSIXlt type

    year.list <- levels(factor(data$Date$year + 1900))

    ### sorting data
    inc.order <- order(data$Date, decreasing = FALSE)
    data <- data[inc.order,]

    ### building an empty data frame
    final.data <- data.frame(data[,1:length(colNameVector)])
    final.data[,] <- NA
    final.data$Date <- as.POSIXlt(final.data$Date)

    year <- substr(data[1,2],1,4)
    origin <- paste(year, "-01-01", sep = "")
    origin <- zoo::as.Date(origin)
    diff <- zoo::as.Date(data[1,2])-origin
    rm(year)

    daySum <- 0
    for(x in year.list)
    {
      daySum <- daySum + Hmisc::yearDays(zoo::as.Date(paste(x, "-01-01", sep = "")))
    }
    daySum <- as.numeric(daySum - diff)

    final.data[1:daySum,] <- NA # remove first few null days because data is not start on 1/1
    final.data[,2] <- seq(data[1,2], by = "1 days", length.out = daySum)

    #### setting rownames
    rownames(data) <- c(1:nrow(data))

    my.index <- match(data$Date, as.POSIXlt(final.data$Date))
    final.data[my.index,] <- data[,]

    for(i in daySum:1)
    {
      if(!is.na(final.data[i,1])){
        tag <- i
        rm(i)
        break;
      }else{
        next;
      }
    }

    final.data <- head(final.data,tag)  # cutting off last empty records

    for(j in 1:nrow(final.data)){
      if(is.na(final.data[j,1])){
        final.data[j,fixedVec] <- data[1,fixedVec]
        final.data[j,pChanged] <- pChangedNum
      }
    }

    final.data$Date <- zoo::as.Date(final.data$Date) # for correcting date time in excel

    colnames(final.data) <- colNameVector

    ### return data
    return(final.data)
  }
