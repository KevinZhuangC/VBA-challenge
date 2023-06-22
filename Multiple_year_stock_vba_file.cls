VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub yearlychange()

    For Each ws In Worksheets
        
        Dim WorksheetNam, ticker_name As String
        Dim Summary_Table_Row, LastRow As Long
        Dim formated_change As String
        Dim ticker_open_price, ticker_close_price, price_change, percent_changet, total_volume As Double
        
        Summary_Table_Row = 2
        ticker_open_price = ws.Cells(2, 3).Value
        total_volume = 0

        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        WorksheetName = ws.Name
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Dim i As Long
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                ticker_name = ws.Cells(i, 1).Value
                ws.Range("I" & Summary_Table_Row).Value = ticker_name
                
                ticker_close_price = ws.Cells(i, 6).Value
                price_change = ticker_close_price - ticker_open_price
                percent_change = price_change / ticker_open_price
                ws.Range("J" & Summary_Table_Row).Value = price_change
                ws.Range("K" & Summary_Table_Row).Value = percent_change
                ws.Range("L" & Summary_Table_Row).Value = total_volume + ws.Cells(i, 7).Value
                
                Summary_Table_Row = Summary_Table_Row + 1
                ticker_open_price = ws.Cells(i + 1, 3).Value
                total_volume = 0
            Else
                total_volume = total_volume + ws.Cells(i, 7).Value
            End If
        
        Next i
        
        Dim min_change, max_change, max_volume As Double
        Dim min_change_rowindex, max_change_rowindex, max_volume_rowindex As Integer
        
        summary_last_row = Summary_Table_Row - 1
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        min_change = Application.WorksheetFunction.Min(ws.Range("K2:K" & summary_last_row))
        max_change = Application.WorksheetFunction.Max(ws.Range("K2:K" & summary_last_row))
        max_volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & summary_last_row))
        ws.Cells(2, 17).Value = max_change
        ws.Cells(3, 17).Value = min_change
        ws.Cells(4, 17).Value = max_volume
        
        max_change_rowindex = Application.WorksheetFunction.Match(ws.Cells(2, 17).Value, ws.Range("K1:K" & summary_last_row), 0)
        min_change_rowindex = Application.WorksheetFunction.Match(ws.Cells(3, 17), ws.Range("K1:K" & summary_last_row), 0)
        max_volume_rowindex = Application.WorksheetFunction.Match(ws.Cells(4, 17).Value, ws.Range("L1:L" & summary_last_row), 0)
        ws.Cells(2, 16).Value = ws.Range("I" & max_change_rowindex).Value
        ws.Cells(3, 16).Value = ws.Range("I" & min_change_rowindex).Value
        ws.Cells(4, 16).Value = ws.Range("I" & max_volume_rowindex).Value
        
        
        For i = 2 To summary_last_row
            If ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0) ' Red
            Else
                ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0) ' Green
            End If
        Next i
        
        ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Next
End Sub
