## Module-2-Challenge

In this module challenge, we have to create a script that loops through all the stocks for each quarter and outputs the required results.

we are provided with the 'multiple_year_stock_data.xlsx' file

[Multiple_year_stock_data.xlsx](https://github.com/user-attachments/files/18312787/Multiple_year_stock_data.xlsx)

we needed to create a macro-enabled excel worksheet which is saved as multiple_year_stock_data.xlsm in this repository

'code.bas' file in the repository contains the code to generate the results.

# How to run the script 

to run the script, open the multiple_year_stock_data.xlsx file and go to the developers tab, click on the visual basic, add a new module with the provided code (which is saved under code.bas) 
<img width="789" alt="Screenshot 2025-01-05 at 8 55 16 PM" src="https://github.com/user-attachments/assets/8e61aeed-7e04-467f-85d3-e4da3f8b763a" />


save the module as a macro enabled excel file and go to the developers tab in the excel sheet and then run macro, 'StockData' 

This will automatically generate the ticker symbol, quarterly change, percentage change and total stock volume.
it also calculate greatest % increase, greatest % decrease and greatest total volume.

some conditional formating is used in the quarterly change column. The negative changes are colored as red and the positive changes as green.no change between the opening and closing price is colored as white.

some number formattings are done in the percentage change column, greatest % increase cell and in the greatest % decrease cell.
    

               
             ws.Cells(cell_count, 11).Value = tickername                       'printing ticker names

   end if

next j

