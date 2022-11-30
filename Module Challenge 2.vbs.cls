VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Alphabet_testing()

For Each ws In Worksheets

Dim Ticker_name As String
Dim Total_Stock As Double
Dim OpenBalance As Double
Dim ClosingBalance As Double
Dim Summary_table_row As Double
Dim YearlyChange As Double
Dim PercentChange As Double


Total_Stock = 0
Summary_table_row = 2

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

OpenBalance = ws.Cells(2, 3).Value

For i = 2 To Lastrow


    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
    Ticker_name = ws.Cells(i, 1).Value
    ws.Range("I" & Summary_table_row).Value = Ticker_name
    
    Total_Stock = Total_Stock + ws.Cells(i, 7).Value
    ws.Range("L" & Summary_table_row).Value = Total_Stock
     
     ClosingBalance = ws.Cells(i, 6).Value
     
     YearlyChange = ClosingBalance - OpenBalance
     ws.Cells(i, 11).NumberFormat = "0.00%"
     PercentChange = ((ClosingBalance - OpenBalance) / OpenBalance) * 100
     
     
     
    ws.Range("J" & Summary_table_row).Value = YearlyChange
    ws.Range("K" & Summary_table_row).Value = PercentChange
     
     Summary_table_row = Summary_table_row + 1
     Total_Stock = 0
     OpenBalance = ws.Cells(i + 1, 3)
     
Else

    Total_Stock = Total_Stock + ws.Cells(i, 7).Value

End If

Next i

 ylastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
            
            
            For j = 2 To ylastrow
            
            If ws.Cells(j, 10).Value > 0 Or ws.Cells(j, 10).Value = 0 Then
            
            ws.Cells(j, 10).Interior.ColorIndex = 4
            
            Else
            
            ws.Cells(j, 10).Interior.ColorIndex = 3
            
            
            End If
            
            
            Next j

Next ws

End Sub






















































