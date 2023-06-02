Sub StockData():

    ' Defining a shorthand for each worksheet in the workbook
    Dim ws As Worksheet
    
    ' Beginning a For Each Loop that allows the Subroutine to run through each sheet consecutively
    For Each ws In Worksheets
    
    ' -------------------------------------------------------------------------------
    ' The following code inputs values for header cells and formats/autofits them along the way
    ' -------------------------------------------------------------------------------
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("I1:L1").EntireColumn.AutoFit

    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("I1:L1").HorizontalAlignment = xlCenter

    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ' Declaring variables needed through the script
    
    Dim tickerRange As Range
    Dim rowCount As Long
    Dim startPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentageChange As Variant
    Dim stockVolume As Double
    Dim cell As Range
    Dim i As Long
    Dim j As Long
    
    ' This helps us find out the total number of rows in Column A, since this can vary based on the year.
    rowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ' We set the range of tickers in Column A using the row count we derived in the previous step
    Set tickerRange = ws.Range("A2:A" & rowCount)
    
    ' Initializing values for the opening price of the ticker as well as stock volume
    startPrice = ws.Cells(2, 3).Value
    stockVolume = 0
    
    ' Initializing values for row variables
    i = 2
    j = 2
    
        ' Beginning a nested For Each Loop that allows us to find the unique ticker values in Column A
    
        For Each cell In tickerRange
        
        ' While we proceed through each row in Column A, we add the values in column G to find total stock volume for a unique ticker as we'll see further down the code
        stockVolume = stockVolume + ws.Cells(i, 7).Value
            
            ' Beginning a next If Conditional that checks if the value in a specific row is NOT equal to the value in the following row. When the condition is met, it would indicate that the ticker moves from one unique value to the next
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            ' Retrieving the unique ticker value and inputing it into Column I
            ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
            
            ' When the above condition is met we know that the current row we are on contains the closing price for that unique ticker
            closePrice = ws.Cells(i, 6).Value
            ' Calculate Yearly Change & Percentage Change knowing that we've initialized the opening price before starting the For Each loop
            yearlyChange = closePrice - startPrice
            percentageChange = (((yearlyChange / startPrice) * 100) & "%")
            
            ' This re-initializes the opening price to the value listed for the next unique ticker
            startPrice = ws.Cells(i + 1, 3).Value
            
            ' Assigning the values we have up until this point that match each unique ticker in Columns J to K
            ws.Cells(j, 10).Value = yearlyChange
            ws.Cells(j, 11).Value = percentageChange
            ws.Cells(j, 12).Value = stockVolume
            
            ' In order to find the total stock volume for the next unique ticker, we have to re-initialize the value back to 0
            stockVolume = 0
            
                ' Beginning another nested conditional that formats the value for Yearly Change
                If yearlyChange < 0 Then
                
                ' If the value is negative, the cell turns red. Otherwise, it will turn green
                ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0)
                
                Else
                
                ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0)
                
                End If
            
            ' Ending the color formatting conditional loop and moving to the next row in Column I
            j = j + 1
            
            End If
        
        ' Ending the loop that finds the sets of unequal values and moves onto the next row in Column A
        i = i + 1
        
        Next cell
    
    ' -------------------------------------------------------------------------------
    ' Now that we've summarized the stock data to give us
    ' 1. The Ticker Symbol (Column I)
    ' 2. Yearly Change with Color Formatting (Column J)
    ' 3. Percentage Change (Column K)
    ' 4. Total Stock Volume for each unique ticker (Column L)
    
    ' The following code will focus on finding out more specific highlights
    ' -------------------------------------------------------------------------------
     
    ' Assigning more variables we need
    Dim newRowCount As Long
    Dim maxIndex As Long
    Dim incIndex As Long
    Dim decIndex As Long
    Dim volIndex As Long
    
    ' Now that we have a new range of unique ticker symbols, we need to find the new amount of rows
    newRowCount = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Using Excel functions to find the maximum and minimum value of Percentage change while concatenating to show the value in a % format
    ws.Range("Q2").Value = (WorksheetFunction.Max(ws.Range("K2:K" & newRowCount)) * 100) & "%"
    ws.Range("Q3").Value = (WorksheetFunction.Min(ws.Range("K2:K" & newRowCount)) * 100) & "%"
    
    ' Using Excel functions to find the maximum value of stock volume
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & newRowCount))
    
    
    ' Knowing the values for maximum percentage change, minimum percentage change and maximum stock volume - we'll used the Match formula to find which row in Columns K & L these values are located in
    ' The secondary lines of code pulls the unique ticker symbols that are offset from the matched rows
    
    incIndex = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & newRowCount), 0)
    ws.Range("P2").Value = ws.Cells((incIndex + 1), 9)

    decIndex = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & newRowCount), 0)
    ws.Range("P3").Value = ws.Cells((decIndex + 1), 9)

    volIndex = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & newRowCount), 0)
    ws.Range("P4").Value = ws.Cells((volIndex + 1), 9)
    
    ' Formatting the remaining columns to fit to the text length
    
    ws.Range("O1:Q1").EntireColumn.AutoFit
    ws.Range("P1:Q1").HorizontalAlignment = xlCenter
  
    ' Ending the For Each Loop for the first sheet and moving on to the next Worksheet in the Workbook to repeat the code
    Next ws

End Sub
