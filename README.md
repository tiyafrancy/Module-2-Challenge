# Module-2-Challenge

In this module challenge, we have to create a script that loops through all the stocks for each quarter and outputs the required results.

we are provided with the multiple_year_stock_data.xlsx file



we needed to creat a macro enabled excel worksheet which is saved as multiple_year_stock_data.xlsm



our sub is called StockData()

to loop through each worksheets, we write the script inside the following

for Each ws In Worksheets


to generate the ticker symbol, we used the following code

For j = 2 To row_count
   
    If ws.Cells(j, 1).Value <> ws.Cells(j + 1, 1) Then
            
             tickername = ws.Cells(j, 1).Value                                       'finding ticker names
               
             ws.Cells(cell_count, 11).Value = tickername                       'printing ticker names

   end if

next j

