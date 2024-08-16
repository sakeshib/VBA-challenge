Sub Alphabetical_Testing()
    Dim ws As Worksheet
    Dim i As Long
    Dim Ticker_Name As String
    Dim Summary_Table As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PriceChange As Double
    Dim PercentChange As Double
    Dim Total_Stock As Double
    Dim lastrow As Long
    Dim QuarterStartRow As Long
    Dim Min As Double
    Dim Max As Double
    Dim Total_Volume As Double
    Dim MinTicker As String
    Dim MaxTicker As String
    Dim VolumeTicker As String
    Dim MinRow As Long
    Dim MaxRow As Long
    Dim VolumeRow As Long
    Dim Condition As Range

    For Each ws In ThisWorkbook.Worksheets

    Summary_Table = 2
    Total_Stock = 0
    QuarterStartRow = 2
    
    'Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Greatest % Increase & Decrease & Total Volume
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    'Ticker Symbol
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_Name = ws.Cells(i, 1).Value
        ws.Range("I" & Summary_Table).Value = Ticker_Name
    
    'Quarterly Change From Opening to Closing Price
        OpenPrice = ws.Cells(QuarterStartRow, 3).Value
        ClosePrice = ws.Cells(i, 6).Value
        PriceChange = ClosePrice - OpenPrice
        ws.Range("J" & Summary_Table).Value = PriceChange

    'Percentage Change From Opening to Closing Price
        If OpenPrice <> 0 Then
            PercentChange = ((PriceChange) / OpenPrice)
        Else
            PercentChange = 0
        End If
            ws.Range("K" & Summary_Table).Value = PercentChange
        
    'Total Stock Volume
        Total_Stock = Total_Stock + ws.Cells(i, 7).Value
        ws.Range("L" & Summary_Table).Value = Total_Stock
        
        Summary_Table = Summary_Table + 1
        
        Total_Stock = 0

        QuarterStartRow = i + 1
        
        Else
            Total_Stock = Total_Stock + ws.Cells(i, 7).Value
    
        End If
        
    Next i
    
    'Find Greatest % Max, Min and Greatest Volume
        Max = WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table - 1))
        Min = WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table - 1))
        Total_Volume = WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table - 1))

    'Find the corresponding rows for Max, Min and Greatest Volume
        MaxRow = ws.Range("K2:K" & Summary_Table - 1).Find(Max).Row
        MinRow = ws.Range("K2:K" & Summary_Table - 1).Find(Min).Row
        VolumeRow = ws.Range("L2:L" & Summary_Table - 1).Find(Total_Volume).Row

    'Grab the Corresponding Ticker
        MaxTicker = ws.Cells(MaxRow, 9).Value
        MinTicker = ws.Cells(MinRow, 9).Value
        VolumeTicker = ws.Cells(VolumeRow, 9).Value

        ws.Range("P2").Value = MaxTicker
        ws.Range("P3").Value = MinTicker
        ws.Range("P4").Value = VolumeTicker
                
        ws.Range("Q2").Value = Max
        ws.Range("Q3").Value = Min
        ws.Range("Q4").Value = Total_Volume
    
    'Formatting
        ws.Columns("J").NumberFormat = "0.00"
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("I:L").AutoFit
        ws.Columns("O:R").AutoFit

        Set Condition = ws.Range("J2:J" & lastrow)

        With Condition.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            .Interior.Color = RGB(255, 0, 0)
        End With

        With Condition.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            .Interior.Color = RGB(0, 255, 0)
        End With

    Next ws
    
End Sub

