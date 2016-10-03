#' Complete the hollow dataset
#'
#' Take time series dataset and fields, then refill the missing date records and other fields.
#' @param inPath A path which is the location of uncompleted dataset which must be xlsx file
#' @param sheet A worksheet name of the dataset
#' @param dateCol.index Date column
#' @param outPath A path where the location of xlsx file of completed dataset should be
#' @param fixedCol.index  A row of column number which should be kept same values with the original
#' @param uninterpolatedCol.index The column number which should be changed to different value into new record.
#' @param uninterpolatedCol.newValue The value of a specific column which should be put into the new record.
#' @import openxlsx zoo Hmisc magrittr
#' @importFrom utils head install.packages
#' @importFrom magrittr %>%
#' @details Real time series sales dataset could be not continuous in 'date' field. e.g., monthly sales data is continuous,
#'  but discrete in daily data.
#'
#'  This hollow dataset is not complete for time series analysis. Function dateRefill.fromFile
#'  is a transformation which tranforms uncomplete dataset into complete dataset.
#' @author Will Kuan
#' @examples # Please refer to the examples of function dateRefill.fromData
#' @export
dateRefill.fromFileToExcel <-
  function(inPath, sheet, dateCol.index, outPath, fixedCol.index, uninterpolatedCol.index, uninterpolatedCol.newValue)
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

    if(!requireNamespace("magrittr",quietly = TRUE)){
      install.packages("magrittr"); requireNamespace("magrittr", quietly = TRUE)
      #stop("Please install package 'magrittr'. ")
    }else{
      requireNamespace("magrittr", quietly = TRUE)
    }

    data <- openxlsx::read.xlsx(inPath, sheet)

    #=============
    data[,dateCol.index] <- zoo::as.Date(data[,dateCol.index], origin = "1899-12-30")
    colNameVector <- colnames(data)

    colnames(data)[dateCol.index] <- "Date"

    data$Date <- as.POSIXlt(data$Date) # transform to POSIXlt type

    year.list <- levels(factor(data$Date$year + 1900))

    ### sorting data
    data <- data[order(data$Date, decreasing = FALSE),]

    ### building an empty data frame
    final.data <- data.frame(data[,1:length(colNameVector)])
    final.data[,] <- NA
    final.data$Date <- as.POSIXlt(final.data$Date)

    year <- substr(data[1, dateCol.index],1,4)
    origin <- paste(year, "-01-01", sep = "")
    origin <- zoo::as.Date(origin)
    diff <- zoo::as.Date(data[1, dateCol.index])-origin

    #=============
    daySum <- sprintf("%s-01-01", year.list) %>%
      zoo::as.Date()

    daySum <- mapply(Hmisc::yearDays, daySum) %>%
      sum()

    daySum <- magrittr::subtract(daySum, diff) %>%
      as.numeric()

    final.data[1:daySum,] <- NA # remove first few null days because data is not start on 1/1
    final.data[, dateCol.index] <- seq(data[1, dateCol.index], by = "1 days", length.out = daySum)

    ### duplicate identical column names
    colnames(data) <- names(final.data)

    ### copy identical record from original data
    my.index <- match(data$Date, as.POSIXlt(final.data$Date))
    final.data[my.index,] <- data[,]

    ### copy fixedCol.value from original data to new records
    final.data <- head(final.data, which(final.data$Date == data$Date[length(data$Date)]))

    #==============
    final.data[which(is.na(final.data[, fixedCol.index[1]])), fixedCol.index] <- data[1, fixedCol.index]
    final.data[which(is.na(final.data[, uninterpolatedCol.index[1]])), uninterpolatedCol.index] <- uninterpolatedCol.newValue

    #==============
    final.data$Date <- zoo::as.Date(final.data$Date) # for correcting date time in excel

    colnames(final.data) <- colNameVector


    ### output data into excel file
    openxlsx::write.xlsx(final.data, file = outPath)

  }

#' Complete the hollow dataset
#'
#' Take time series dataset and fields, then refill the missing date records and other fields.
#' @param data The data.frame dataset which is ready to be processed
#' @param dateCol.index Date column
#' @param fixedCol.index  A row of column number which should be kept same values with the original
#' @param uninterpolatedCol.index The column number which should be changed to different value into new record.
#' @param uninterpolatedCol.newValue The value of a specific column which should be put into the new record.
#' @import zoo Hmisc magrittr
#' @importFrom utils head install.packages
#' @importFrom magrittr %>%
#' @return The dataset which is completed.
#' @details Real time series sales dataset could be not continuous in 'date' field. e.g., monthly sales data is continuous,
#'  but discrete in daily data.
#'
#'  This hollow dataset is not complete for time series analysis. Function dateRefill.fromFile
#'  is a transformation which tranforms uncomplete dataset into complete dataset.
#' @author Will Kuan
#' @examples # mydata <- data.example
#' # mydata.final <- dateRefill.fromData(data = mydata,dateCol = 2,fixedVec = c(3:10),
#' #                                     uninterpolatedCol.index = 11,uninterpolatedCol.newValue = 0)
#' @export
dateRefill.fromData <-
  function(data, dateCol.index, fixedCol.index, uninterpolatedCol.index, uninterpolatedCol.newValue)
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

    if(!requireNamespace("magrittr",quietly = TRUE)){
      install.packages("magrittr"); requireNamespace("magrittr", quietly = TRUE)
      #stop("Please install package 'magrittr'. ")
    }else{
      requireNamespace("magrittr", quietly = TRUE)
    }

    data <- data.frame(data)

    #=============
    data[,dateCol.index] <- zoo::as.Date(data[,dateCol.index], origin = "1899-12-30")
    colNameVector <- colnames(data)

    colnames(data)[dateCol.index] <- "Date"

    data$Date <- as.POSIXlt(data$Date) # transform to POSIXlt type

    year.list <- levels(factor(data$Date$year + 1900))

    ### sorting data
    data <- data[order(data$Date, decreasing = FALSE),]

    ### building an empty data frame
    final.data <- data.frame(data[,1:length(colNameVector)])
    final.data[,] <- NA
    final.data$Date <- as.POSIXlt(final.data$Date)

    year <- substr(data[1, dateCol.index],1,4)
    origin <- paste(year, "-01-01", sep = "")
    origin <- zoo::as.Date(origin)
    diff <- zoo::as.Date(data[1, dateCol.index])-origin

    #=============
    daySum <- sprintf("%s-01-01", year.list) %>%
      zoo::as.Date()

    daySum <- mapply(Hmisc::yearDays, daySum) %>%
      sum()

    daySum <- magrittr::subtract(daySum, diff) %>%
      as.numeric()

    final.data[1:daySum,] <- NA # remove first few null days because data is not start on 1/1
    final.data[, dateCol.index] <- seq(data[1, dateCol.index], by = "1 days", length.out = daySum)

    ### duplicate identical column names
    colnames(data) <- names(final.data)

    ### copy identical record from original data
    my.index <- match(data$Date, as.POSIXlt(final.data$Date))
    final.data[my.index,] <- data[,]

    ### copy fixedCol.value from original data to new records
    final.data <- head(final.data, which(final.data$Date == data$Date[length(data$Date)]))

    #==============
    final.data[which(is.na(final.data[, fixedCol.index[1]])), fixedCol.index] <- data[1, fixedCol.index]
    final.data[which(is.na(final.data[, uninterpolatedCol.index[1]])), uninterpolatedCol.index] <- uninterpolatedCol.newValue

    #==============
    final.data$Date <- zoo::as.Date(final.data$Date) # for correcting date time in excel

    colnames(final.data) <- colNameVector

    ### return data
    return(final.data)
  }
