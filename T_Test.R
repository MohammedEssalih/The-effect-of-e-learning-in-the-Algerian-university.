# Load the readxl library to read Excel files
library(readxl)

# Read the Excel file located at the specified path and store it in Sliman_Data
Sliman_Data <- read_excel("C:/Users/utente/Downloads/Sliman Data.xlsx")

# View the loaded data in a spreadsheet-like viewer
View(Sliman_Data)

# Perform t-tests on various variables grouped by 'test' and store the results
result_total <- t.test(total ~ test, data = Sliman_Data)
result1 <- t.test(unity ~ test, data = Sliman_Data)
result2 <- t.test(coherence ~ test, data = Sliman_Data)
result3 <- t.test(grammar ~ test, data = Sliman_Data)
result4 <- t.test(spelling ~ test, data = Sliman_Data)
result5 <- t.test(vocabulary ~ test, data = Sliman_Data)

# Install and load the 'officer' and 'flextable' packages if they are not already installed
if (!require("officer")) {
  install.packages("officer")
}
if (!require("flextable")) {
  install.packages("flextable")
}
library(officer)
library(flextable)

# Create a data frame from the t-test results (example shown for one t-test result)
t_test_df <- data.frame(
  Statistic = round(result_total$statistic, 3),
  Parameter = round(result_total$parameter, 3),
  P.value = format.pval(result_total$p.value, digits = 3),
  Conf.Int.Lower = round(result_total$conf.int[1], 3),
  Conf.Int.Upper = round(result_total$conf.int[2], 3),
  Estimate = round(result_total$estimate, 3),
  Null.Value = round(result_total$null.value, 3),
  Alternative = result_total$alternative,
  Method = result_total$method,
  Data.Name = result_total$data.name
)

# Create a flextable object to format the t-test results data frame
ft <- flextable(t_test_df)

# Create a new Word document
doc <- read_docx()

# Add a title to the Word document
doc <- body_add_par(doc, "T-test Results", style = "heading 1")

# Add the formatted t-test results table to the Word document
doc <- body_add_flextable(doc, ft)

# Save the Word document with the specified filename
print(doc, target = "t_test_result.docx")
