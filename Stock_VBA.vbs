Sub Testing()
Dim Header As Boolean
Header = False

For Each ws In Worksheets
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    Dim Ticker As String
    Dim Dates As String
    Dim Open_price As Double
    Open_price = 0
    Dim Close_price As Double
    Close_price = 0
    Dim Yearly_change As Double
    Yearly_change = 0
    Dim Percent_change As Double
    Percent_change = 0
    Dim Ticker_Volume As Double
    Ticker_Volume = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    If Header Then
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    Else
        Header = True
    End If
    Open_price = ws.Cells(2, 3).Value
    
        For i = 2 To LastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                Close_price = ws.Cells(i, 6).Value
                Yearly_change = Close_price - Open_price
                Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
                
                If Open_price <> 0 Then
                    Percent_change = (Yearly_change / Open_price) * 100
                Else
                    ws.Range("J" & Summary_Table_Row).Value = "Error"
                End If
                
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = Yearly_change
                If (Yearly_change > 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Yearly_change <= 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_change) & "%")
                ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume
                Summary_Table_Row = Summary_Table_Row + 1
                Ticker_Volume = 0
                Close_price = 0
                Open_price = ws.Cells(i + 1, 3).Value
             Else
                Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
            End If
        Next i
Next ws
End Sub
