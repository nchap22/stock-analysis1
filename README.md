# stock-analysis1
## Overview of Analysis
Steve is interested in expanding the dataset and he wants to look at the entire stock market over the last couple of years (2017 and 2018).Our job is to refactor the code and to amke the code more efficient by having fewer steps and using less memorey. This way when Steve is looking up a dataset that has thousands of stocks versus only a dozen stocks, it's still running efficently and isn't taking super long to load. 
## Results 
#### Using images and examples of your code, compare the stock performance between 2018 and 2018, as well as the execution times of the original script and the refactored script.

Below I have the original code and the refactored code and I have the time that it took to load each. I found that the original code took longer to load than the refactored code for both years. The 2017 original code showed that it took 0.609 seconds and the refactored 2017 took 0.109 seconds. The 2018 original code showed that it took 0.6132 and the refactored code showed that it took 0.109 seconds. 

## Original Code

       Sub AllStocksAnalysis()
       
      '1) Format the output sheet on All Stocks Analysis worksheet
      
      Worksheets("All Stocks Analysis").Activate
      
       Range("A1").Value = "All Stocks (2018)"
       
       'Create a header row
       
       Cells(3, 1).Value = "Ticker"
       
       Cells(3, 2).Value = "Total Daily Volume"
       
       Cells(3, 3).Value = "Return"
       
        '2) Initialize array of all tickers
        
        Dim tickers(11) As String
        
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
        
       '3a) Initialize variables for starting price and ending price
       
         Dim startingPrice As Single
         
        Dim endingPrice As Single
        
         '3b) Activate data worksheet
         
         Worksheets("2018").Activate
         
        '3c) Get the number of rows to loop over
        
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
       '4) Loop through tickers
          For i = 0 To 11
          
         ticker = tickers(i)
         
         TotalVolume = 0
         
         '5) loop through rows in the data
         
          Worksheets("2018").Activate
          
          For j = 2 To RowCount
          
           '5a) Get total volume for current ticker
           
           If Cells(j, 1).Value = ticker Then
           
               TotalVolume = TotalVolume + Cells(j, 8).Value
               
           End If
           
           '5b) get starting price for current ticker
           
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
           
               startingPrice = Cells(j, 6).Value
               
           End If
           
           '5c) get ending price for current ticker
           
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
           
              endingPrice = Cells(j, 6).Value
              
           End If
           
       Next j
       
       '6) Output data for current ticker
       
       Worksheets("All Stocks Analysis").Activate
       
       Cells(4 + i, 1).Value = ticker
       
       Cells(4 + i, 2).Value = TotalVolume
       
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
       
         Next i
         
      'Formatting
      
    Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.FontStyle = "Bold"
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#,##0"
    
    Range("C4:C15").NumberFormat = "0.0%"
    
    Columns("B").AutoFit
    
    dataRowStart = 4
    
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then 
        
            Cells(i, 3).Interior.Color = vbGreen   
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed   
            
        End If   
        
    Next i
    
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub

## Original Code Time 2018
<img width="401" alt="OriginalCode2018_TimeAmount" src="https://user-images.githubusercontent.com/110268006/189543984-fc8aec33-cf5e-4082-8318-b9321e63f80b.png">

## Original Code Time 2017
<img width="550" alt="OriginalCode2017_TimeAmount" src="https://user-images.githubusercontent.com/110268006/189544010-b858aa13-41cf-4565-83fa-ecddbcbd9f84.png">

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

  '2a) Create a for loop to initialize the tickerVolumes to zero.
  
    For i = 0 To 11
    
tickerVolumes(i) = 0

Next i

    '2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        'If  Then
        
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
    End If
    
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
        
         'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
         
        'If  Then
        
           If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
           
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            '3d Increase the tickerIndex.
            
             tickerIndex = tickerIndex + 1
             
             End If
             
             Next i
             
        'End If   
        
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
For i = 0 To 11

    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = tickers(i)
    
    Cells(4 + i, 2).Value = tickerVolumes(i)
    
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
       Next i

    'Formatting
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A3:C3").Font.FontStyle = "Bold"
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#,##0"
    
    Range("C4:C15").NumberFormat = "0.0%"
    
    Columns("B").AutoFit
    
    dataRowStart = 4
    
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
        
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If  
        
    Next i
    
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub

## Refactored Time Amount 2017
<img width="575" alt="RefactoredCode2017_TimeAmount" src="https://user-images.githubusercontent.com/110268006/189544176-3df9298f-9fda-4e34-b3c3-ff0ea1a11f19.png">

## Refactored Time Amount 2018
<img width="593" alt="RefactoredCode2018_TimeAmount" src="https://user-images.githubusercontent.com/110268006/189544181-fbdfc3b1-451c-48bd-b932-9db1aa65e4a5.png">

## Code Outputs


### Code Output for 2017
<img width="433" alt="2017 Code" src="https://user-images.githubusercontent.com/110268006/189544546-d306d7f7-02b6-4932-a8d1-14561e59876c.png">

### Refactored Code Output for 2018
<img width="352" alt="2018 refactored code" src="https://user-images.githubusercontent.com/110268006/189544599-0098d94e-ec40-4f2a-bf19-df42165a071a.png">

### Original Code Output for 2018

<img width="1043" alt="2018 Original Code Output" src="https://user-images.githubusercontent.com/110268006/189544608-5264fc39-13f8-40ee-8c1d-5371898a4260.png">

## Summary: In a summary statement, address the following questions.

### What are the advantages or disadvantages of refactoring code?

#### Advantages
1. Code is more efficent and takes less time to execute
2. Code is cleaner and easier to update
3. Refactring helps find bugs in advance and can improve code design
#### Disadvantages
1. Could introduce new errors and bugs
2. Could cost more money
3. Could also take more time to go through and refactor code

### How do these pros and cons apply to refactoring the original VBA script?
#### Pro
1. The efficency did increase when I refactored the code because the time it took to load icreased with the vba compared ot the original script.
#### Con
1. It does take more time trying to refactor the code and in this challenge it took some time to understand the vba syntax for refactoring. 

