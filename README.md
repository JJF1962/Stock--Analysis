# Stock--Analysis
##  Overview Visual Basic For Aplications Project
The project consist in to edit or refactor a code and analysis of the years 2017 and 2018 Stock Market data consolidated in the workbook VBA_Challenge.xlsm. Loop through all the data and present a writting analysis providing the findings, and demostrate that the code is more efficient in terms using less memory, improve code logic, and finally make it easier for future readers.
##  Results
The challenge is very specific in the way that the delivery must be sumbmitted, detailed as follows.
### Refactor VBA Code and Measure Performance
The code was refactored is fully described in the Microsoft Visual Basic Editor of the file VBA_Challenge, following the four callenges, First it was created a tickerIndex variable and set it equal to zero was set to zero befor itinerated over all the roads, for i = 0 to 11, ticker index = tickers (i), the number 11, is the result of the summatory of all the stocks, from zero to 11, also it was created three output arrays as indicated in the challenge, as variable, we use Dim. Dim tickersVolume As Long  and  Dim tickerStartingPrices As Single and  Dim tickerEndingtingPrices As Single, aditionally the refactor clode we wrote acript to increase the curret tickerVolumes variable and add the ticker volume to the current stock ticker, and we add a conditional as well as detailed as follows For j = 2 to RowCount, it is commun use j but we can substitute if needed, we made ticker volumes = o For j = 2 to RowCount and as mentioned the conditional if Cells(j, 1).Value = thicker index then, with the finality, that in case the next row's ticker not match, increase the tickerIndex. Furthermore we st the stored values  from ticketStarting Prices and the Ticker ending prices using If and End if,and add codes for formatting the cells, and set the year button and the clear button thet was more challenge and requeried a Code, all was established following the Canvas, class excersize aand instructions, VBA challenge and the original code provided with the challenge.
### The Stock Analysis
Please see below the 2 charts with the results for 2017 and 2018, respectivelly.
  ![This is an image](https://github.com/JJF1962/Stock--Analysis/blob/main/Capture%20Results%20%26%20Enable%20Time%202017%20Refactor%20Analysis.PNG)
  ![This is an image](https://github.com/JJF1962/Stock--Analysis/blob/main/Capture%20Results%20%26%20Enable%20Time%202018%20Refactor%20Analysis.PNG)

The Tickers have a better retuns in 2017 than 2018, showing a positive results  with the exception of TERP with minus 7.2%, However, that was not the case in the year 2018, were ten Tickers show negative results , and only two Tickers ENPH and Run shows positive results, looks like is very risky concentrate al the investment in DQ, that have a great performance  in 2017 with a a return of 199.4%, but in 2018 has anegative return of - 62.6%.
The refactored code ran for 2017 in 1.53 secs and in 2018 in 1.43 secs, as shown in the previos pictures, however the original code ran as shown in the figures below
  ![This is an image] ()
  ![this is an image](https://github.com/JJF1962/Stock--Analysis/blob/main/Capture%20Elapse%20Run%20time%20original%20code%202017%20final.PNG)
  
##  Summary
