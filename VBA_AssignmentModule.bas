Attribute VB_Name = "Module2"
Sub CalculateQuarterlyChangeAndPercentageAllSheets()
    Dim ws As Worksheet
    Dim wsNames As Variant
    Dim lastRow As Long, i As Long
    Dim ticker As String, CurrentTicker As String
    Dim openingPrice As Double, closingPrice As Double
    Dim Total_Volume As Double
    Dim outputRow As Long
    
    wsNames = Array("Q1", "Q2", "Q3", "Q4")
    For Each wsName In wsNames
        Set ws = ThisWorkbook.Sheets(wsName)
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' headers for the results in columns I, J, K, L
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Initializing output row
        outputRow = 2
        ' Initialize variables
        CurrentTicker = ws.Cells(2, 1).Value
        openingPrice = ws.Cells(2, 3).Value
        Total_Volume = 0
               
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            If ticker <> CurrentTicker Then
                closingPrice = ws.Cells(i - 1, 6).Value
                ws.Cells(outputRow, 9).Value = CurrentTicker
                ws.Cells(outputRow, 10).Value = closingPrice - openingPrice
                ws.Cells(outputRow, 12).Value = Total_Volume
                
                If openingPrice <> 0 Then ' added this step to avoid dividing 0
                    ws.Cells(outputRow, 11).Value = Format(((closingPrice - openingPrice) / openingPrice) * 100, "0.00") & "%"
                Else
                    ws.Cells(outputRow, 11).Value = "N/A"
                End If
                               
                outputRow = outputRow + 1
                
                ' Reseting variables for the next ticker
                CurrentTicker = ticker
                openingPrice = ws.Cells(i, 3).Value
                Total_Volume = ws.Cells(i, 7).Value
            Else
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            End If
        Next i
     Next wsName
End Sub
Sub GreatestIncreaseAndDecrease()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim maxChange As Double
    Dim minChange As Double
    Dim maxVolume As Double
    Dim tickerMaxChange As String
    Dim tickerMinChange As String
    Dim tickerMaxVolume As String
    
    ' Initializing variables
    maxChange = 0
    minChange = 0
    maxVolume = 0
    
    For Each ws In ThisWorkbook.Worksheets(Array("Q1", "Q2", "Q3", "Q4"))
        
        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        For i = 2 To lastRow
            If ws.Cells(i, "K").Value > maxChange Then
                maxChange = ws.Cells(i, "K").Value
                tickerMaxChange = ws.Cells(i, "I").Value
            End If
            If ws.Cells(i, "K").Value < minChange Then
                minChange = ws.Cells(i, "K").Value
                tickerMinChange = ws.Cells(i, "I").Value
            End If
            If ws.Cells(i, "L").Value > maxVolume Then
                maxVolume = ws.Cells(i, "L").Value
                tickerMaxVolume = ws.Cells(i, "I").Value
            End If
        Next i
        
        ' Storing results in columns O, P, Q
        ws.Cells(2, "Q").Value = maxChange
        ws.Cells(2, "P").Value = tickerMaxChange
        ws.Cells(3, "Q").Value = minChange
        ws.Cells(3, "P").Value = tickerMinChange
        ws.Cells(4, "Q").Value = maxVolume
        ws.Cells(4, "P").Value = tickerMaxVolume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(, 17).Value = "Value"
        
    Next ws

End Sub
Sub HighlightQuarterlyChange()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    For Each ws In ThisWorkbook.Worksheets(Array("Q1", "Q2", "Q3", "Q4"))
        
        lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        For i = 1 To lastRow
            If ws.Cells(i, "J").Value > 0 Then
                ws.Cells(i, "J").Interior.ColorIndex = 4
            ElseIf ws.Cells(i, "J").Value < 0 Then
                ws.Cells(i, "J").Interior.ColorIndex = 3
            End If
        Next i
    Next ws
End Sub
Sub HighlightPercentageChange()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    For Each ws In ThisWorkbook.Worksheets(Array("Q1", "Q2", "Q3", "Q4"))
        
        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        For i = 1 To lastRow
            If ws.Cells(i, "K").Value > 0 Then
                ws.Cells(i, "k").Interior.ColorIndex = 4
            ElseIf ws.Cells(i, "K").Value < 0 Then
                ws.Cells(i, "K").Interior.ColorIndex = 3
            End If
        Next i
    Next ws
End Sub

