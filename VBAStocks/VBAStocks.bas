Attribute VB_Name = "Module4"
Sub StockAnalysis()
'WITH EACH WORKSHEET IN WORKBOOK: Create new table for summary and second table for extremes

For Each ws In Worksheets
'FOR ONE WORKSHEET, comment out the above line and the next to last line ('Next ws')
'AND comment in the next 2 lines
'Dim ws As Worksheet
'Set ws = ActiveSheet

    'Sort Stock Data by Stock Ticker and Date, Date is in YYYYMMDD format which will sort ok
    ws.Columns("A:G").Sort key1:=ws.Range("A2"), order1:=xlAscending, _
                    key2:=ws.Range("B2"), order2:=xlAscending, _
                    Header:=xlYes

    'Find the last row of data
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create analysis table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    Dim Volume As Double
        Volume = 0
    Dim LastClose As Double
    Dim FirstOpen As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TableRow As Integer
        TableRow = 1
    Dim i As Long
    Dim Ticker As String
    
    Ticker = " "
    
    For i = 2 To LastRow
    
        'For each cell add up the volume, it will reset at ticker change
        Volume = CLng(ws.Cells(i, 7).Value) + Volume
        
        'If the cell has a non-zero value for the open price
        If ws.Cells(i, 3).Value > 0 And ws.Cells(i, 1).Value <> Ticker Then
                'Define and hold the opening value for the new ticker
                FirstOpen = ws.Cells(i, 3).Value
                'And change the ticker value
                Ticker = ws.Cells(i, 1).Value
                'Debug.Print (Ticker)
                'Debug.Print (FirstOpen)
        End If
            
            'For the last row before the ticker symbol changes
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value _
                And ws.Cells(i, 1).Value = Ticker Then
                'And ws.Cells(i + 1, 10).Value <> "" _
                'Start entering in the next row of the analysis table
                TableRow = TableRow + 1
                'Put in the ticker symbol before the change
                ws.Cells(TableRow, 9).Value = Ticker
                'Put in the volume accrued
                ws.Cells(TableRow, 12) = Volume
                'Grab the value at close on the last day
                LastClose = ws.Cells(i, 6).Value
                'Subtract the value at open on the first day
                YearlyChange = LastClose - FirstOpen
                'Place this value in the table
                ws.Cells(TableRow, 10).Value = YearlyChange
                'Calculate the Percent
                 PercentChange = YearlyChange / FirstOpen
                'Place Percent in the table
                ws.Cells(TableRow, 11).Value = PercentChange
                ws.Cells(TableRow, 11).NumberFormat = "0.00%"
                'Format the Yearly change
                    If ws.Cells(TableRow, 10).Value > 0 Then
                        ws.Cells(TableRow, 10).Interior.ColorIndex = 10
                    ElseIf ws.Cells(TableRow, 10).Value < 0 Then
                        ws.Cells(TableRow, 10).Interior.ColorIndex = 3
                    End If
                'Reset the volume to 0 for the next ticker
                Volume = 0
            End If
    Next i

'NEXT TABLE WITH INCREASE / DECREASE / VOLUME

    'Find the last row of the summary table
    Dim LastRow2 As Long
    LastRow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    'Define the values that will go into the table
    Dim ChgMax As Double
        ChgMax = -1
    Dim ChgMin As Double
        ChgMin = 0
    Dim VolMax As Double
        VolMax = 0
    Dim ChgMaxTicker As String
    Dim ChgMinTicker As String
    Dim VolMaxTicker As String
    
    'Look for Max, Min, Vol extremes in each row
    For i = 2 To LastRow2
        
            If ws.Cells(i, 11).Value > ChgMax Then
                ChgMax = ws.Cells(i, 11).Value
            End If
            If ws.Cells(i, 11).Value < ChgMin Then
                ChgMin = ws.Cells(i, 11).Value
            End If
            If ws.Cells(i, 12).Value > VolMax Then
                VolMax = ws.Cells(i, 12).Value
            End If
            
    Next i
    
    'Use the values found in the previous block to find ticker
    For i = 2 To LastRow2
    
        If ws.Cells(i, 11).Value = ChgMax Then
            ChgMaxTicker = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 11).Value = ChgMin Then
            ChgMinTicker = ws.Cells(i, 9).Value
        End If
        
        If ws.Cells(i, 12).Value = VolMax Then
            VolMaxTicker = ws.Cells(i, 9).Value
        End If
    
    Next i

    'Create second table
    'Need ticker and value for each stat
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    'Fill in title and values for Increase
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("P2").Value = ChgMaxTicker
    ws.Range("Q2").Value = ChgMax
    ws.Range("Q2").NumberFormat = "0.00%"
    'Fill in title and values for Decrease
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("P3").Value = ChgMinTicker
    ws.Range("Q3").Value = ChgMin
    ws.Range("Q3").NumberFormat = "0.00%"
    'Fill in title and values for Volume
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P4").Value = VolMaxTicker
    ws.Range("Q4").Value = VolMax
    
    ws.Columns("I:Q").AutoFit
    
Next ws

End Sub
