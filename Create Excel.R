create.workbook <- function(Mth) {
# create new workbook
wb <- createWorkbook()

#--------------------
# DEFINE CELL STYLES 
#--------------------
# title and subtitle styles
title_style <- CellStyle(wb) + 
  Font(wb, heightInPoints = 16,
       # color = "black", 
       isBold = TRUE, 
       underline = 1)

subtitle_style <- CellStyle(wb) + 
  Font(wb, heightInPoints = 14,
       isItalic = TRUE,
       isBold = FALSE)

# data table styles
rowname_style <- CellStyle(wb) +
  Font(wb, isBold = TRUE)

cs2 <- CellStyle(wb) + Border(color="black", position=c("BOTTOM", "LEFT","TOP","RIGHT"), pen=c("BORDER_THIN", "BORDER_THIN","BORDER_THIN","BORDER_THIN"))

colname_style <- CellStyle(wb) +
  Font(wb, isBold = TRUE) +
  Alignment(wrapText = TRUE, horizontal = "ALIGN_CENTER") +
  Border(color = "black",
         position = c("TOP", "BOTTOM","LEFT","RIGHT"),
         pen = c("BORDER_THICK", "BORDER_THICK","BORDER_THIN","BORDER_THIN"))

all_Colstyle<- rep(list(cs2), dim(All.Data)[2]) 
names(all_Colstyle) <- seq(1, dim(All.Data)[2], by = 1)
#-------------------------
# CREATE & EDIT WORKSHEET 
#-------------------------
# create worksheet
Cars <- createSheet(wb, sheetName = "Cars")

# helper function to add titles
xlsx.addTitle <- function(sheet, rowIndex, title, titleStyle) {
  rows <- createRow(sheet, rowIndex = rowIndex)
  sheetTitle <- createCell(rows, colIndex = 1)
  setCellValue(sheetTitle[[1,1]], title)
  setCellStyle(sheetTitle[[1,1]], titleStyle)
}

# add title and sub title to worksheet
xlsx.addTitle(sheet = Cars, rowIndex = 1, 
              title = "1974 Motor Trend Car Data",
              titleStyle = title_style)

xlsx.addTitle(sheet = Cars, rowIndex = 2, 
              title = "Performance and design attributes of 32 automobiles",
              titleStyle = subtitle_style)

# add data frame to worksheet
addDataFrame(All.Data, sheet = Cars, startRow = 3, startColumn = 1,
             colnamesStyle = colname_style,
             rownamesStyle = rowname_style,
             colStyle = all_Colstyle,
             row.names = FALSE)

# change row name column width
setColumnWidth(sheet = Cars, colIndex = 1, colWidth = 18)
autoSizeColumn(sheet = Cars, colIndex=2:ncol(All.Data))
# save workbook
saveWorkbook(wb, file = paste("Report Mastercard MCAE On ",Mth,".xlsx",sep = ""))
}