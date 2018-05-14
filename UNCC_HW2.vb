Sub calcStockHistory()
    'This sub will run through each sheet in the notebook, skipping the first line
    'It will total the volume and capture the ticker signal
    
    
    '----------------------
    'Part I: variable setup and initialization
    '----------------------
    Dim currentCell As Integer
    Dim wsCount As Integer
    Dim ws As Worksheet
    Dim currentTicker, nextTicker, greatestIncTicker, greatestDecTicker, greatestVolTicker As String
    Dim i, j, totalVolume, percentChange, outputTracker, currentTickerOpen, currentTickerClose, greatestInc, greatestDec, greatestVol As Double
    Dim headers, headersColumn As Variant
    
    Dim volumeCol, openCol, closeCol, tickerOutputCol, yearChangeCol, percentChangeCol, totalStockVolCol As Integer
        
    'static variables for column offsets
    volumeCol = 7
    openCol = 3
    closeCol = 6
    tickerOutputCol = 9
    yearChangeCol = 10
    percentChangeCol = 11
    totalStockVolCol = 12
    
    'initialize variables
    greatestInc = 0
    greatestVol = 0
    greatestDec = 0
    greatestIncTicker = ""
    greatestDecTicker = ""
    greatestVolTicker = ""
    
    'initiate headers
    headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "", "", "Ticker", "Value")
    headersColumn = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
    
    
    'get count of sheets in workbook
    wsCount = ActiveWorkbook.Worksheets.Count

    '-------------------------
    'Part II: Create summary table for each ticker in workbook
    '-------------------------
        
    'iterate over all worksheets
    For i = 1 To wsCount
        'create data headers for current sheet ticker summary and greatest increase, decrease and volume for year
        ActiveWorkbook.Worksheets(i).Range("I1:P1") = headers
        ActiveWorkbook.Worksheets(i).Range("N2:N4") = headersColumn
        
    
        'set counter for first ticker in sheet
        j = 2
        
        'Tracks row of output in current sheet for the ticker summary
        outputTracker = 2
        
        'set active ticker symbol
        currentTicker = ActiveWorkbook.Worksheets(i).Cells(j, 1)
        nextTicker = ActiveWorkbook.Worksheets(i).Cells(j + 1, 1)
        currentTickerOpen = ActiveWorkbook.Worksheets(i).Cells(j, openCol)
        
        'Check to make sure we are not at the end of the list on the current sheet
        Do While currentTicker <> ""
            If currentTicker <> nextTicker Then
                'we are the last instance of the ticker on this sheet, grab it's closing data
                currentTickerClose = ActiveWorkbook.Worksheets(i).Cells(j - 1, closeCol)
                
                'output current ticker symbol
                ActiveWorkbook.Worksheets(i).Cells(outputTracker, tickerOutputCol).Value = currentTicker
                'calculate and fill yearly change
                ActiveWorkbook.Worksheets(i).Cells(outputTracker, yearChangeCol).Value = currentTickerClose - currentTickerOpen
            
                'determine if yearly change is positive or negative and color appropriately
                If currentTickerClose - currentTickerOpen >= 0 Then
                    ActiveWorkbook.Worksheets(i).Cells(outputTracker, yearChangeCol).Interior.Color = RGB(0, 255, 0)
                Else
                    ActiveWorkbook.Worksheets(i).Cells(outputTracker, yearChangeCol).Interior.Color = RGB(255, 0, 0)
                End If
                
                'Calculate the percent change from open to close
                'Make sure currentTickerOpen is not 0, else you will get a divide by 0 error
                If currentTickerOpen <> 0 Then
                    percentChange = ((currentTickerClose - currentTickerOpen) / currentTickerOpen)
                Else
                    percentChange = 0
                End If
                
                'write percent change to the ticker summary
                ActiveWorkbook.Worksheets(i).Cells(outputTracker, percentChangeCol).Value = percentChange
                ActiveWorkbook.Worksheets(i).Cells(outputTracker, percentChangeCol).NumberFormat = "0.00%"
                
                'check if current ticker's percent increase is the largest seen so far
                If percentChange >= greatestInc Then
                    'store the percent increase and the stocks ticker
                    greatestIncTicker = currentTicker
                    greatestInc = percentChange
                End If
                
                'check if the current ticker's percent decrease is the largest seen so far
                If percentChange < greatestDec Then
                    'store the percent decrease and the stock's ticker
                    greatestDecTicker = currentTicker
                    greatestDec = percentChange
                End If
                
                'write ticker's total volume to the worksheet
                ActiveWorkbook.Worksheets(i).Cells(outputTracker, totalStockVolCol).Value = totalVolume
                
                'check if the current ticker's traded cvolume is the largest seen so far
                If totalVolume > greatestVol Then
                    'store the percent decrease and the stock's ticker
                    greatestVolTicker = currentTicker
                    greatestVol = totalVolume
                End If
                
                'reset for next ticker
                totalVolume = 0
                currentTickerOpen = ActiveWorkbook.Worksheets(i).Cells(j + 1, openCol)
                'increment to next line of output
                outputTracker = outputTracker + 1
            Else
                'next ticker is same as current so add volume to total
                totalVolume = totalVolume + ActiveWorkbook.Worksheets(i).Cells(j, 7)
            End If
            
            'increment to next line in workbook,
            j = j + 1
            currentTicker = nextTicker
            nextTicker = ActiveWorkbook.Worksheets(i).Cells(j, 1)
        Loop
        
        'Once the loop is complete all of the greatest variables will contain the correct ticker and corresponding value, output to excel sheet
        ActiveWorkbook.Worksheets(i).Cells(2, 15).Value = greatestIncTicker
        ActiveWorkbook.Worksheets(i).Cells(2, 16).Value = greatestInc
        ActiveWorkbook.Worksheets(i).Cells(2, 16).NumberFormat = "0.00%"
        
        ActiveWorkbook.Worksheets(i).Cells(3, 15).Value = greatestDecTicker
        ActiveWorkbook.Worksheets(i).Cells(3, 16).Value = greatestDec
        ActiveWorkbook.Worksheets(i).Cells(3, 16).NumberFormat = "0.00%"
        
        ActiveWorkbook.Worksheets(i).Cells(4, 15).Value = greatestVolTicker
        ActiveWorkbook.Worksheets(i).Cells(4, 16).Value = greatestVol
        
        're-initialize variables for next sheet
        greatestInc = 0
        greatestVol = 0
        greatestDec = 0
        greatestIncTicker = ""
        greatestDecTicker = ""
        greatestVolTicker = ""
        
    Next
    
    
    
End Sub

Sub reset()
     'get count of sheets in workbook
    wsCount = ActiveWorkbook.Worksheets.Count
    
    For j = 1 To wsCount
        'this sub is meant to reset the output columns, useful for testing
        For i = 1 To 1000
            ActiveWorkbook.Worksheets(j).Cells(i, 9) = ""
            ActiveWorkbook.Worksheets(j).Cells(i, 10) = ""
            ActiveWorkbook.Worksheets(j).Cells(i, 10).Interior.ColorIndex = 0
            ActiveWorkbook.Worksheets(j).Cells(i, 11) = ""
            ActiveWorkbook.Worksheets(j).Cells(i, 12) = ""
            ActiveWorkbook.Worksheets(j).Cells(i, 14) = ""
            ActiveWorkbook.Worksheets(j).Cells(i, 15) = ""
            ActiveWorkbook.Worksheets(j).Cells(i, 16) = ""
        Next i
    Next j
End Sub

