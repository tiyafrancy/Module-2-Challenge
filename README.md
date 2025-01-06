### Module-2-Challenge

In this module challenge, we have to create a script that loops through all the stocks for each quarter and outputs the required results.

we are provided with the 'multiple_year_stock_data.xlsx' file

[Multiple_year_stock_data.xlsx](https://github.com/user-attachments/files/18312787/Multiple_year_stock_data.xlsx)

we needed to create a macro-enabled excel worksheet which is saved as multiple_year_stock_data.xlsm in this repository

'code.bas' file in the repository contains the code to generate the results.

# How to run the script 

to run the script, open the multiple_year_stock_data.xlsx file and go to the developers tab, click on the visual basic 


<img width="769" alt="Screenshot 2025-01-05 at 9 07 25 PM" src="https://github.com/user-attachments/assets/33c8a5e3-e9dd-40c4-8117-f09478315bfc" />


add a new module with the provided code (which is saved under code.bas) 


<img width="1370" alt="Screenshot 2025-01-05 at 9 05 44 PM" src="https://github.com/user-attachments/assets/05adc60c-ba4a-4954-975f-e38e6cba46de" />


save the module as a macro enabled excel file and go to the developers tab in the excel sheet and then run macro, 'StockData' 


<img width="432" alt="Screenshot 2025-01-05 at 9 09 23 PM" src="https://github.com/user-attachments/assets/fcb5bdd6-76d2-4f5e-9046-8488a23b1c87" />


This will automatically generate the ticker symbol, quarterly change, percentage change and total stock volume.
it also calculate greatest % increase, greatest % decrease and greatest total volume.

some conditional formating is used in the quarterly change column. The negative changes are colored as red and the positive changes as green. If there is no change then the color is white.

some number formattings are also done in the percentage change column, greatest % increase cell and in the greatest % decrease cell.

I have also done some formatting to the date column just to match with the required result. 

We have 4 sheets in this single workbook. We can loop through each sheets by using this single macro sub function and produce same outputs in each sheets. 

# Acknowledgment 

I have done this assignment using the help of some internet searches and some help from my SMU Instructor.

