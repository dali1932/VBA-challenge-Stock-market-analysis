Attribute VB_Name = "Module1"
Sub Stock_analysis()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    Dim ticker As String
    Dim ttlvol As Double
    ttlvol = 0
    
    Dim summary_row As Integer
    summary_row = 2
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            ttlvol = ttlvol + ws.Cells(i, 7).Value
    
            ws.Range("I" & summary_row).Value = ticker
            ws.Range("L" & summary_row).Value = ttlvol
            
            summary_row = summary_row + 1
            ttlvol = 0 ' Reset ttlvol for the next ticker
        Else
            ttlvol = ttlvol + ws.Cells(i, 7).Value
        End If
    Next i
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Price Change"
    ws.Range("L1").Value = "Total Stock Value"
    Next ws
End Sub

Sub Stock_pricechg()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
Dim current_ticker As String
Dim y_o, y_c, yrchg, ptchg As Double
y_o = 0
y_c = 0
Dim r As Integer
r = 2

y_o = ws.Cells(2, 3).Value
current_ticker = ws.Cells(2, 1).Value
For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
If ws.Cells(i + 1, 1) <> current_ticker Or ws.Cells(i, 1) = ws.Cells(Rows.Count, 1).End(xlUp).Row Then
y_c = ws.Cells(i, 6).Value

yrchg = y_c - y_o
ptchg = y_c / y_o - 1
ws.Range("J" & r).Value = yrchg
ws.Range("K" & r).Value = ptchg
ws.Range("J" & r).Style = "Currency"
ws.Range("K" & r).NumberFormat = "0.00%"
r = r + 1
'reset variables for the new ticker
current_ticker = ws.Cells(i + 1, 1).Value
y_o = ws.Cells(i + 1, 3).Value
yrchg = 0
ptchg = 0
End If

If ws.Range("J" & i).Value < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
ElseIf ws.Range("J" & i).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4

End If

Next i


Next ws
End Sub

Sub Greatest()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
Dim i As Integer
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest total volume"
ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Range("K2:" & "K" & ws.Cells(ws.Rows.Count, 11).Row))
ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Range("K2:" & "K" & ws.Cells(ws.Rows.Count, 11).Row))
ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("L2:" & "L" & ws.Cells(ws.Rows.Count, 11).Row))
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"

For i = 2 To Cells(Rows.Count, 11).End(xlUp).Row
If ws.Cells(i, 11).Value = ws.Range("Q2").Value Then
ws.Range("P2").Value = ws.Cells(i, 9).Value

ElseIf ws.Cells(i, 11).Value = ws.Range("Q3").Value Then
ws.Range("P3").Value = ws.Cells(i, 9).Value

ElseIf ws.Cells(i, 12).Value = ws.Range("Q4").Value Then
ws.Range("P4").Value = ws.Cells(i, 9).Value
End If
Next i
Next ws
End Sub

Sub all_sub():
 Stock_analysis
 Stock_pricechg
 Greatest
 
End Sub


