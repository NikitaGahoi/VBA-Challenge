# VBA-Challenge
This folder contains 2 files:
1) The .vbs script is named  "VBA_Challenge_Nikita.vbs"
2) Screenshots of the analysis for each year named as "Screenshots of the VBA challenge_Nikita.pptx"

For the analysis following steps were followed:
Step 1: Find the last row of the sheet
Step 2: To create a script that loops through the stock data and summarizes it:
  (i)Respective colums with headers were created
  (ii) A loop was run to check the ticker symbol for each cell, and the open_value of the stock is assigned as the first open_value
    -if the ticker symbol for the next row is same as the ticker symbol of the next row, the total stock volume was added for thst stock
    -if the ticker symbol for the next row is NOT same  as the ticker symbol of the next row, then ticker symbol data was added into the summary table, the         
    close_value for the stock is assigned, and total stock volume was updated by adding the stock volume of the last cell found with the same name
    -calculated the Yearly change by substracting the open_value to close_value (close_value-open_value) and %v change was claculates using 
     (yearly_change/open_value)*100
    -color_coded the +ve and -ve values for positive and negative yearly change and %change
    -Assign a new_open value and ticker_name to the variable
    -Reset the value of the total stock volume to 0
Step 3: Find the last row of the summary table
Step 4: To create a script that loops through the stock data to find the greatest % increase and decrease and to find the maximum total stock volume
  (i) Respective colums with headers were created
  (ii) three variables Max_PR, Min_PR and GTV (greatest total volume) were created anfd their values were set to 0
  (ii) Three IF statements were created 
        -to check maximum % change 
        -to check minimum % change 
        -to check greatest total volume 
Step 5: once the code ran successfully on one sheet of alphabetical_testing
Step 6 : A loop was created to run the code on all the worksheet
Step 7: Once that code worked on the alphabetical_testing
Step 8: Same code was modified and ran on the main workbook containg yearly data for the stocks
