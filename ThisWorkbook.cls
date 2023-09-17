VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub StockTracker()

    Dim ws As Worksheet
    Dim Tickers As Collection
    Dim LastRow As Long, i As Long, OutputRow As Long
    Dim TickerValue As String, OpenPrice As Double, ClosePrice As Double
    Dim PriceDifference As Double, PercentageChange As Double, TSV As Double
    Dim MaxIncrease As Double, MaxDecrease As Double, MaxVolume As Double
    Dim TickerMaxIncrease As String, TickerMaxDecrease As String, TickerMaxVolume As String
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "2018" Or ws.Name = "2019" Or ws.Name = "2020" Then

            Set Tickers = New Collection

            MaxIncrease = -999999
            MaxDecrease = 999999
            MaxVolume = 0
            TickerMaxIncrease = ""
            TickerMaxDecrease = ""
            TickerMaxVolume = ""

            LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            OutputRow = 3

            For i = 2 To LastRow
                TickerValue = ws.Cells(i, 1).Value
                
                On Error Resume Next
                Tickers.Add TickerValue, CStr(TickerValue)
                On Error GoTo 0
                
                If Err.Number = 0 Then
                    OpenPrice = ws.Cells(i, 3).Value
                    TSV = 0

                    ' Loop to find the last occurrence for the close price and summing up the TSV
                    While i < LastRow And ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value
                        TSV = TSV + ws.Cells(i, 7).Value
                        i = i + 1
                    Wend
                    
                    ' Add the volume of the last occurrence of the ticker to TSV
                    TSV = TSV + ws.Cells(i, 7).Value
                    
                    ClosePrice = ws.Cells(i, 6).Value

                    ' Calculate Yearly Change
                    PriceDifference = ClosePrice - OpenPrice
                    If OpenPrice <> 0 Then
                        PercentageChange = (PriceDifference / OpenPrice)
                    Else
                        PercentageChange = 0
                    End If

                    ' Check for Max and Min values
                    If PercentageChange > MaxIncrease Then
                        MaxIncrease = PercentageChange
                        TickerMaxIncrease = TickerValue
                        
                    End If
                    
                    If PercentageChange < MaxDecrease Then
                        MaxDecrease = PercentageChange
                        TickerMaxDecrease = TickerValue
                    End If
                    
                    If TSV > MaxVolume Then
                        MaxVolume = TSV
                        TickerMaxVolume = TickerValue
                    End If

                    ' Output results on the current worksheet
                    ws.Cells(OutputRow, 16).Value = TickerValue
                    ws.Cells(OutputRow, 17).Value = PriceDifference
                    ws.Cells(OutputRow, 18).Value = PercentageChange
                    ws.Cells(OutputRow, 19).Value = TSV

                    OutputRow = OutputRow + 1
                Else
                    Err.Clear
                End If
            Next i

            ' Output summary data on the current worksheet
            ws.Cells(4, 22).Value = "Ticker with Max % Increase:"
            ws.Cells(4, 23).Value = TickerMaxIncrease
            ws.Cells(4, 24).Value = MaxIncrease
            
            ws.Cells(5, 22).Value = "Ticker with Max % Decrease:"
            ws.Cells(5, 23).Value = TickerMaxDecrease
            ws.Cells(5, 24).Value = MaxDecrease
            
            ws.Cells(6, 22).Value = "Ticker with Max Volume:"
            ws.Cells(6, 23).Value = TickerMaxVolume
            ws.Cells(6, 24).Value = MaxVolume
        End If
    Next ws

End Sub

