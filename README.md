# Stock-Analysis
Stock data analysis
## Overview of Project
Refactor previous version of stock analysis visual basic for applications code to improve efficiency. 

### Purpose
The purpose of this project was to take a previously created Excel Macro Enabled workbook that was used to gather stock information for a client and improve the existing VBA code to improve the efficieny of the workbook.    

## Analysis and Challenges
I began by taking the existing code and removing the existing loop structure while leaving the code that dealt with the formatting.  I then began working through the looping structure to improve the efficiency of the previously written code.  The steps I went through for the refactoring are as follows:

1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i

2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If

3c) check if the current row is the last row with the selected ticker. 
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If

3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

 4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
    Next i
### Results
The speed for the original file for 2017:
![Original 2017 Speed](https://github.com/john10roberts/stock-analysis/blob/main/Resources/Green_stocks_ProcessingTime_2017.png)

The speed for the refactored file for 2017:
![Refactored 2017 Speed](https://github.com/john10roberts/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

The speed for the original file for 2018
![Original 2018 Speed](https://github.com/john10roberts/stock-analysis/blob/main/Resources/Green_stocks_ProcessingTime_2018.png)

The speed for the refactored file for 2018
![Refactored 2018 Speed](https://github.com/john10roberts/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

- What are the advantages or disadvantages of refactoring code?
"Refactoring is a controlled technique for improving the design of an existing code base." (Fowler, n.d.).  The major advantages of code refactoring are improving the efficiency of the existing code.  This has been demonstrated in our refactoring of the source data as the speeds for 2017 and 2018 have clearly increased.  The major disadvantage to refactoring is you are taking code that’s working as expected in its current format and changing it.  This could cause the code to stop working altogether or introduce other issues that might not be tested for correctly.  Additionally, the amount of time spent refactoring the code is like the amount of time it might have taken to just develop the solution from scratch.  

- How do these pros and cons apply to refactoring the original VBA?
For this stock project we were clearly able to improve the efficiency of the code and make the entire project execute quicker.  If we were dealing with a much larger excel sheet this increase in efficiency could be a tremendous pick up.  However, the increase in efficiency of the existing spreadsheet is probably not worth the effort in refactoring the code.  

References
Fowler, M. (n.d.). Refactoring. Retrieved from martinFowler.com: https://martinfowler.com/books/refactoring.html
