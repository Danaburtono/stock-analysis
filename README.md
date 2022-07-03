# stock-analysis

# Overview of Project

## Purpose

The purpose of this analysis is to refactor a Microsoft Excel VBA code to collect certain stock options. I aimed to make the code more applicable for a larger volume of stock options and make the code more efficient.


## Results

The data that is presented includes two datasets with stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the stock volume. The goal is to retrieve the ticker, the total daily volume, and the return on each stock. The stock performance between 2017 and 2018 was dramatically different. In 2017(left chart), 11 of 12 tickers had net positive returns on their stock. In 2018(right chart), nearly all but 2 tickers had negative returns. Changes implemented to the code increased code running time from 0.656sec to 0.136sec for the 2017 and in 2018, 0.625sec to 0.125sec.

![Screen Shot 2022-07-01 at 5 34 20 PM](https://user-images.githubusercontent.com/107026442/176980111-07ac2a6a-15b9-4388-96ef-1db9389c3421.png)
![Screen Shot 2022-07-01 at 5 34 03 PM](https://user-images.githubusercontent.com/107026442/176980115-143d92f7-895e-420f-b647-ac654368c3ca.png)

![VBA_Challenge_2017_original](https://user-images.githubusercontent.com/107026442/176979466-022e4393-6895-4770-a648-db5838f87b54.png)
![VNA_Challenege_2017](https://user-images.githubusercontent.com/107026442/176979471-412d11d9-1a31-463a-bb5d-db88a735d15f.png)

![VBA_Challenge_2018_orginal](https://user-images.githubusercontent.com/107026442/176979474-95b9c9ea-eee5-45cf-a01d-4cb1eb6b3a1f.png)
![VBA_Challlenge_2018](https://user-images.githubusercontent.com/107026442/176979478-e0cdf82e-2566-4c3b-9444-792908649dc4.png)


### The Refractored Code

    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1


## Summary

Typically new code isn't as efficient as it could be because the programmer is troubleshooting a new workflow. Refactoring helps clean up and organize code. A few advantages of a cleaner code are there are fewer possibilities for bugs, unnecessary syntax, faster processing time, and increased readability for future programmers. Refactoring creates a collaborative coding environment because refactoring is viewed as a fundamental part of programming. You will likely be able to find someone willing to refactor your code on [stackoverflow](https://stackoverflow.com). The downside of allowing your code to be refactored is releasing it to the public; people may or may not take your scripts for their own purposes. 

The pros of refactoring the original VBA script are run times have improved and can now be used to analyze thousands of data points. The cons to refactoring in VBA is that many people don't trust VBA and it can lead to viruses and other malicious events because VBA is a less trusted software it is harder to get people to implement and share Macros.
