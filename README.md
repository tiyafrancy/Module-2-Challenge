## Module-2-Challenge

In this module challenge, we have to create a script that loops through all the stocks for each quarter and outputs the required results.

we are provided with the multiple_year_stock_data.xlsx file

[Multiple_year_stock_data.xlsx](https://github.com/user-attachments/files/18312787/Multiple_year_stock_data.xlsx)

we needed to creat a macro enabled excel worksheet which is saved as multiple_year_stock_data.xlsm

This is the code that we used


[UploadingAttribute VB_Name = "Module1"
Sub StockData()

For Each ws In Worksheets               'loops through each worksheets
     
    Dim row_count As Long

        row_count = ws.Range("A" & Rows.Count).End(xlUp).Row          ' stores last active row count
   
       
    Dim j As Long
    
    Dim tickername As String
    
    Dim cell_count As Integer
    cell_count = 2
    
    Dim total_stock_volume As Double
    total_stock_volume = 0
                
    Dim opening_price As Double
    opening_price = 0
    
    Dim closing_price As Double
    
    Dim quarterly_change As Double
    
    Dim percentage_change As Double
    
    Dim greatest_percentage_increase As Double
    greatest_percentage_increase = 0
    
    Dim greatest_percentage_decrease As Double
    greatest_percentage_decrease = 0
    
    Dim greatest_total_volume As Double
    greatest_total_volume = 0
    
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_total_ticker As String
    
    ' Adding Headers to new output columns
    
    ws.Cells(1, 11).Value = "Ticker"
    ws.Cells(1, 12).Value = "Quarterly Change"
    ws.Cells(1, 13).Value = "Percentage Change"
    ws.Cells(1, 14).Value = "Total Stock Volume"
    
     
     ' loop to process each row
     
        For j = 2 To row_count
        
         total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value         'finding total stock volume
         
         If opening_price = 0 Then
         
            opening_price = ws.Cells(j, 3).Value            'storing opening price for each ticker
        
         End If
        
          If ws.Cells(j, 1).Value <> ws.Cells(j + 1, 1) Then
            
             tickername = ws.Cells(j, 1).Value                                       'finding ticker names
               
             ws.Cells(cell_count, 11).Value = tickername                       'printing ticker name
             
             ws.Cells(cell_count, 14).Value = total_stock_volume            'printing total stock volume
                
               
                
                closing_price = ws.Cells(j, 6).Value
                
                quarterly_change = closing_price - opening_price               'calculating quarterly change for each ticker
                
                ws.Cells(cell_count, 12).Value = quarterly_change                  'printing quarterly change value
                
                percentage_change = quarterly_change / opening_price       'calculating percentage change
                
                ws.Cells(cell_count, 13).Value = percentage_change          'printing percentage change
                
                ws.Cells(cell_count, 13).NumberFormat = "0.00%"             'formating the percentage change column
                
                
                'Conditional formatting for positive/ negative changes
                
                If quarterly_change < 0 Then
                
                    ws.Cells(cell_count, 12).Interior.Color = RGB(255, 0, 0)           'Red Interior color
                
                ElseIf quarterly_change > 0 Then
                
                    ws.Cells(cell_count, 12).Interior.Color = RGB(0, 255, 0)             'Green Interior color
                
                Else
                
                    ws.Cells(cell_count, 12).Interior.Color = RGB(255, 255, 255)       'White Interior color
                
                End If
                
                'Calculating and storing the greatest percentage increase and its corresponding ticker name
                
                If percentage_change > greatest_percentage_increase Then
                
                    greatest_percentage_increase = percentage_change
                    greatest_increase_ticker = tickername
                         
                End If
                
                'Calculating and storing the greatest percentage decrease and its corresponding ticker name
                
                If percentage_change < greatest_percentage_decrease Then
                
                    greatest_percentage_decrease = percentage_change
                    greatest_decrease_ticker = tickername
                    
                End If
                
                'Calculating and storing the total stock volume and its corresponding ticker name
                
                If total_stock_volume > greatest_total_volume Then
                
                    greatest_total_volume = total_stock_volume
                    greatest_total_ticker = tickername
                    
                End If
                
                'Printing the greatest percentage increase, greatest percentage decrease and total stock volume
                
                ws.Cells(1, 18).Value = "Ticker"
                ws.Cells(1, 19).Value = "Value"
                
                ws.Cells(2, 17).Value = " Greatest % Increase"
                ws.Cells(2, 18).Value = greatest_increase_ticker
                ws.Cells(2, 19).Value = greatest_percentage_increase
                ws.Cells(2, 19).NumberFormat = "0.00%"
                
                ws.Cells(3, 17).Value = " Greatest % Decrease"
                ws.Cells(3, 18).Value = greatest_decrease_ticker
                ws.Cells(3, 19).Value = greatest_percentage_decrease
                ws.Cells(3, 19).NumberFormat = "0.00%"
                
                ws.Cells(4, 17).Value = "Greatest Total Volume"
                ws.Cells(4, 18).Value = greatest_total_ticker
                ws.Cells(4, 19).Value = greatest_total_volume
                
                'resetting the variables
                
                opening_price = 0
                 
                total_stock_volume = 0
                
                cell_count = cell_count + 1
          
                
            End If
            
     Next j
                            
Next ws


End Sub

 code.bas…]()



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

