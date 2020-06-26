Attribute VB_Name = "Module1"
Sub MYSD()
'declaring some variables: open_price, closing_price, ticker, summary_table_index
Dim last_row As Long
Dim Ticker As String
Dim summary_table_index As Integer
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim total_volume As Double
Dim pct_change As Double
Dim ws As Worksheet

'For loop to cycle through workbook
For Each ws In Worksheets

ws.Activate

'find the last row
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    summary_table_index = 2
    total_volume = 0
'start at row 2 make room for header
'start first open_price and ticker
    open_price = Cells(2, 3).Value
    Ticker = Cells(2, 1).Value
'iterate through the rows
    For I = 2 To last_row
        close_price = Cells(I, 6).Value
        total_volume = total_volume + Cells(I, 7)
    
    ' need to check ticker change
    If Ticker <> Cells(I + 1, 1).Value Then
    yearly_change = close_price - open_price
    
    If open_price <> 0 Then
    pct_change = (yearly_change / open_price)
        Else
        pct_change = 100
        End If
        ' updating the table
        Range("I" & summary_table_index).Value = Ticker
        Range("J" & summary_table_index).Value = yearly_change
        'color conditional for Yearly Change
        If yearly_change >= 0 Then
            Range("J" & summary_table_index).Interior.ColorIndex = 4
        Else
            Range("J" & summary_table_index).Interior.ColorIndex = 3
        End If
        'format percentage
        Range("K" & summary_table_index).Value = FormatPercent(pct_change, 0)
        Range("L" & summary_table_index).Value = total_volume
        'move summary table down next row
        summary_table_index = summary_table_index + 1
       'reset the open_price and ticker for next
        open_price = Cells(I + 1, 3).Value
        Ticker = Cells(I + 1, 1).Value
        total_volume = 0
     End If
     
     Next I
     
     'Name Headers and format columns
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        ActiveSheet.Range("I1:L1").Columns.AutoFit
       
        
     
    
        Next
        
        
End Sub
