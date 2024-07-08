Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()
    Dim ws As Worksheet
    Dim outputWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim volume As Double
    Dim quarterChange As Double
    Dim percentageChange As Double
    Dim outputRow As Long
    Dim totalVolume As Double
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim maxPercentIncreaseTicker As String
    Dim maxPercentDecreaseTicker As String
    Dim maxTotalVolumeTicker As String
    
    ' Initialize max values
    maxPercentIncrease = 0
    maxPercentDecrease = 0
    maxTotalVolume = 0
    
    ' Create a new worksheet for output
    Set outputWs = ThisWorkbook.Sheets.Add
    outputWs.Name = "Quarterly Analysis"
    
    ' Set the headers for the output sheet
    outputWs.Cells(1, 1).Value = "Ticker"
    outputWs.Cells(1, 2).Value = "Quarterly Change"
    outputWs.Cells(1, 3).Value = "Percent Change"
    outputWs.Cells(1, 4).Value = "Total Stock Volume"
    outputWs.Cells(9, 10).Value = "Greatest % Increase"
    outputWs.Cells(9, 11).Value = "Greatest % Decrease"
    outputWs.Cells(9, 12).Value = "Greatest Total Volume"
    
    outputRow = 2
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Quarterly Analysis" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ' Create dictionary to store data
            Dim tickerData As Object
            Set tickerData = CreateObject("Scripting.Dictionary")
            
            ' Collect data from each row
            For i = 2 To lastRow
                ticker = ws.Cells(i, 1).Value
                openingPrice = ws.Cells(i, 3).Value
                closingPrice = ws.Cells(i, 6).Value
                volume = ws.Cells(i, 7).Value
                
                If Not tickerData.Exists(ticker) Then
                    tickerData.Add ticker, Array(openingPrice, closingPrice, volume)
                Else
                    Dim x As Variant
                    x = tickerData(ticker)
                    x(1) = closingPrice
                    x(2) = x(2) + volume
                    tickerData(ticker) = x
                End If
            Next i
            
            ' Get keys from the dictionary
            Dim tickerKeys As Variant
            tickerKeys = tickerData.Keys
            
            ' Write summary data to the current worksheet and calculate metrics
            For j = LBound(tickerKeys) To UBound(tickerKeys)
                ticker = tickerKeys(j)
                Dim store As Variant
                store = tickerData(ticker)
                openingPrice = store(0)
                closingPrice = store(1)
                totalVolume = store(2)
                quarterChange = closingPrice - openingPrice
                If openingPrice <> 0 Then
                    percentageChange = (quarterChange / openingPrice) * 100
                Else
                    percentageChange = 0
                End If
                
                ' Output the results
                outputWs.Cells(outputRow, 1).Value = ticker
                outputWs.Cells(outputRow, 2).Value = quarterChange
                outputWs.Cells(outputRow, 3).Value = percentageChange
                outputWs.Cells(outputRow, 4).Value = totalVolume
                
                ' Apply conditional formatting
                If quarterChange >= 0 Then
                    outputWs.Cells(outputRow, 2).Interior.Color = RGB(0, 255, 0) ' Green
                    outputWs.Cells(outputRow, 3).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    outputWs.Cells(outputRow, 2).Interior.Color = RGB(255, 0, 0) ' Red
                    outputWs.Cells(outputRow, 3).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                ' Update max values
                If percentageChange > maxPercentIncrease Then
                    maxPercentIncrease = percentageChange
                    maxPercentIncreaseTicker = ticker
                End If
                If percentageChange < maxPercentDecrease Then
                    maxPercentDecrease = percentageChange
                    maxPercentDecreaseTicker = ticker
                End If
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    maxTotalVolumeTicker = ticker
                End If
                
                outputRow = outputRow + 1
                
                
            Next j
        End If
    Next ws
    
    ' Output the max values
    outputWs.Cells(10, 10).Value = maxPercentIncreaseTicker & " (" & maxPercentIncrease & "%)"
    outputWs.Cells(10, 11).Value = maxPercentDecreaseTicker & " (" & maxPercentDecrease & "%)"
    outputWs.Cells(10, 12).Value = maxTotalVolumeTicker & " (" & maxTotalVolume & ")"
    
    'creating quarterly Analysis into 4 equal parts in Q1,Q2, Q3 and Q4.
     
     Dim wsQ1 As Worksheet, wsQ2 As Worksheet, wsQ3 As Worksheet, wsQ4 As Worksheet
    
    Dim rowsPerSheet As Long
    

    ' Define the number of rows per quarter
    rowsPerSheet = 1500
    
    ' Set the source worksheet
    Set ws = ThisWorkbook.Sheets("Quarterly Analysis")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Create or clear the destination worksheets
    On Error Resume Next
    Set wsQ1 = ThisWorkbook.Sheets("Q1")
    Set wsQ2 = ThisWorkbook.Sheets("Q2")
    Set wsQ3 = ThisWorkbook.Sheets("Q3")
    Set wsQ4 = ThisWorkbook.Sheets("Q4")
    
    If wsQ1 Is Nothing Then
        Set wsQ1 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsQ1.Name = "Q1"
    Else
        wsQ1.Cells.Clear
    End If
    
    If wsQ2 Is Nothing Then
        Set wsQ2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsQ2.Name = "Q2"
    Else
        wsQ2.Cells.Clear
    End If
    
    If wsQ3 Is Nothing Then
        Set wsQ3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsQ3.Name = "Q3"
    Else
        wsQ3.Cells.Clear
    End If
    
    If wsQ4 Is Nothing Then
        Set wsQ4 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsQ4.Name = "Q4"
    Else
        wsQ4.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Copy headers to each quarter sheet
    ws.Rows(1).Copy Destination:=wsQ1.Rows(1)
    ws.Rows(1).Copy Destination:=wsQ2.Rows(1)
    ws.Rows(1).Copy Destination:=wsQ3.Rows(1)
    ws.Rows(1).Copy Destination:=wsQ4.Rows(1)
    
    ' Copy data to each quarter sheet
    For i = 1 To 4
        For j = 1 To rowsPerSheet
            ws.Rows((i - 1) * rowsPerSheet + j + 1).Copy Destination:=ThisWorkbook.Sheets("Q" & i).Rows(j + 1)
        Next j
    Next i
    
    ' Optional: Autofit columns in each quarter sheet
    wsQ1.Columns.AutoFit
    wsQ2.Columns.AutoFit
    wsQ3.Columns.AutoFit
    wsQ4.Columns.AutoFit

End Sub

    


