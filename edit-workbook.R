library(openxlsx)

download.file(url = "http://publications.naturalengland.org.uk/file/5985083561607168",
              destfile = "bng-calculator.xlsm")

bng_wb <- openxlsx::loadWorkbook(file = "bng-calculator.xlsm")

habitat <- "Woodland and forest - Lowland mixed deciduous woodland"
area <- 2

# Saving the workbook
openxlsx::saveWorkbook(bng_wb, file = "no-change.xlsm")

# Enter some data
openxlsx::writeData(wb = bng_wb, sheet = "A-1 Site Habitat Baseline", startCol = 4, startRow = 11, x = habitat)
openxlsx::writeData(wb = bng_wb, sheet = "A-1 Site Habitat Baseline", startCol = 5, startRow = 11, x = area)

# Open the workbook & navigate to the above sheet
# you won't be able to change any of the value where you could previously
openxlsx::saveWorkbook(bng_wb, file = "entered-data.xlsm")


