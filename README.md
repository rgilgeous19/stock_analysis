# **ANALYSIS OF RESULTS**


## Overview of Project

    The purpose of this project is to analyze stock data, using VBA, for Steve's 
    parents, so they would be able to make a sensible decision of 
    which stocks to invest in. His parents initially wanted to invest in
    one stock, but Steve advised them not to put all their eggs in 
    one basket. This project finds the total daily
    volume and yearly return for each stock for the years 2017 and 
    2018, so Steve's parents would be able to choose the stocks that would
    give them the maximum return.
    
    
## Results

- ### ***2017 Analysis***
            
         In 2017, DQ stock had the highest return rate of 199.45% with a total
         daily volume of 35,796,200. However, TERP stock had the lowest
         return rate of -7.21% with a total daily volume of 139,402,800.
         From this information, we can infer that the price of the DQ stock ha 
         been rapidly increasing over the year, while the price of TERP had
         been on the decline. 
  
- ### ***2018 Analysis***

          In 2018, most of the stocks had a negative return rate except for 
          ENPH with a return rate of 81.92% and RUN with a return rate of 83.95%.
          Since most of the stocks were in the "red", we can assume that there
          might have been something external, like inflation or current world events,
          affecting the stock market.
          
          
          
  Based on the stock data analyzed in 2017 and 2018, I would advise Steve's 
  parents to invest in ENPH, RUN and SEDG. ENPH and RUN stocks had positive 
  returns in both years, and specifically in 2018 when all other stocks were
  on the decline. This shows the resilience from the external impacts of the 
  stock market. They should also invest in SEDG because they had a great return 
  of 184.5% in 2017. Even though they had -7.8% return in 2018, the negative 
  return rate is small compared to the return rate in 2017.
          
          
      
 ## Refactored Code
 
        
    Sub AllStocksAnalysisRefactored()
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
        tickerIndex = 0
        
        
    '1b) Create three output arrays
        
        Dim tickerVolumes(12) As Long
        
        Dim tickerStartingPrices(12) As Single
        
        Dim tickerEndingPrices(12) As Single
        
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
        For I = 0 To 11
        
            tickerVolumes(I) = 0
            
            tickerStartingPrices(I) = 0
            
            tickerEndingPrices(I) = 0
            
            
        Next I
        
        
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    For I = 2 To RowCount
    
        '3a) Increase volume for current ticker
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(I, 8).Value
            
            
        
     '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
          If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I - 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerStartingPrices(tickerIndex) = Cells(I, 6).Value
                
            
            End If
            
            
     'End If
        
        
    '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            
            
            If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
            
                    tickerEndingPrices(tickerIndex) = Cells(I, 6).Value
                    
            
            End If
            


      '3d Increase the tickerIndex.
      
            
            If Cells(I, 1).Value = tickers(tickerIndex) And Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
                    
                    tickerIndex = tickerIndex + 1
                    
                    
             End If
             
            
        'End If
    
    Next I
    
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For I = 0 To 11
    
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + I, 1).Value = tickers(I)
        
        Cells(4 + I, 2).Value = tickerVolumes(I)
        
        Cells(4 + I, 3).Value = tickerEndingPrices(I) / tickerStartingPrices(I) - 1
        
        
    Next I
    
    'Formatting
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("A3:C15").HorizontalAlignment = xlCenter
    Range("C4:C15").NumberFormat = "0.00%"
    Range("A3:C3").Font.Italic = True
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For I = dataRowStart To dataRowEnd
        
        If Cells(I, 3) > 0 Then
            
            Cells(I, 3).Interior.COLOR = vbGreen
            
        Else
        
            Cells(I, 3).Interior.COLOR = vbRed
            
        End If
        
        
    Next I
 
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub
    
    
  ### Short Analysis of Code
        
        In 2017, the original VBA code ran in 0.6953125 seconds, while the
        refactored code ran in 0.125 seconds.                          
        
![VBA_Challenge_2017](https://user-images.githubusercontent.com/103302566/164868123-dfffe487-8498-4ba1-be25-a3cc1a58386d.png)
      
         In 2018, the original VBA code ran in 0.703125 seconds, while the 
         refactored code ran in 0.125 seconds.
         
![VBA_Challenge_2018](https://user-images.githubusercontent.com/103302566/164868159-dd1decb5-f787-4b44-ac25-0efaf1838149.png)

       
## Summary
    
    - In general, some advantages of refactoring code would be that it helps clean up the original code, 
      making it more efficient and saving time. It also makes the original code more comprehensible and maintable. 
      However, refactoring code may introduce new errors and bugs into the code. It also may not transfer everything
      you needed from your original code to make it successful.
      
    - The biggest advantage of refactoring my code was the decrease in the time it took the code to run. 
      In the original VBA script, my code was more freehanded and I made more mental notes of
      the steps I needed to take. However, the refactored code provided more details on the steps 
      I needed to take in order to run the code efficiently. On the original VBA script, I was able to 
      make the code my own, but the refactored code did not copy everything from my original code, 
      specifically the formatting. Overall, refactoring the orginal code provided more advantages than disadvantages for me.

