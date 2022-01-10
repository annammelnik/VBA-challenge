VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub test2()

Dim x As Integer
LastCell = Cells(Rows.Count, 1).End(xlUp).Row
Set TickerNumber = Range(Cells(2, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, 1))
x = 2
Dim TickerTotal As Double
TickerTotal = 0
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Annual Change in Price"
Cells(1, 12).Value = "Annual Percent Change"
Cells(1, 13).Value = "Total Volume"

For i = 2 To LastCell
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        Cells(x, 10).Value = Cells(i, 1).Value
        TickerAmount = Application.WorksheetFunction.CountIf(TickerNumber, Cells(i, 1).Value)
        BegPrice = Cells(i, 3).Value
        EndPrice = Cells(i + TickerAmount - 1, 6).Value
        ChangePrice = EndPrice - BegPrice
        Cells(x, 11).Value = ChangePrice
        If ChangePrice > 0 Then
            Cells(x, 11).Interior.ColorIndex = 10
        ElseIf ChangePrice < 0 Then
            Cells(x, 11).Interior.ColorIndex = 9
        End If
        PercChange = (EndPrice - BegPrice) / BegPrice
        Cells(x, 12).Value = PercChange
        Set TickerNumberVol = Range(Cells(i, 7), Cells(i + TickerAmount - 1, 7))
        Cells(x, 13).Value = Application.WorksheetFunction.Sum(TickerNumberVol)
        x = x + 1
    End If
    
    
Next i


End Sub

