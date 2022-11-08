# VBA-challenge
Sub demo()

'Different Worksheets
Dim a As Integer

a = Application.Worksheets.Count

For j = 1 To a
Worksheets(j).Activate

'Each Worksheet

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Declaring variables
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Vol As Double
Dim Pchange As Double
Dim OpenP As Double
Dim CloseP As Double
Dim i As Long


Total_Vol = 0
Yearly_Change = 0
Percent_Change = 0

Dim summary As Integer
summary = 2

'Naming the headers
ActiveSheet.Range("j1").Value = "Ticker"
ActiveSheet.Range("k1").Value = "Yearly Change"
ActiveSheet.Range("l1").Value = "Percent Change"
ActiveSheet.Range("m1").Value = "Total Stock Volume"

OpenP = Cells(2, 3).Value

'Conditional
For i = 2 To lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value
Total_Vol = Total_Vol + Cells(i, 7).Value
CloseP = Cells(i, 6).Value

Yearly_Change = (CloseP - OpenP)
Cells(summary, 11).NumberFormat = "0.00"

Percent_Change = (Yearly_Change / OpenP)
Cells(summary, 12).NumberFormat = "0.00%"

Range("J" & summary).Value = Ticker
Range("K" & summary).Value = Yearly_Change
Range("L" & summary).Value = Percent_Change
Range("M" & summary).Value = Total_Vol


summary = summary + 1
OpenP = Cells(i + 1, 3)
Total_Vol = 0

Else

Total_Vol = Total_Vol + Cells(i, 7).Value

End If

Next i

lastrow2 = Cells(Rows.Count, 10).End(xlUp).Row

For p = 2 To lastrow2
    If Cells(p, 11).Value >= 0 Then
        Cells(p, 11).Interior.ColorIndex = 4
    ElseIf Cells(p, 11).Value < 0 Then
        Cells(p, 11).Interior.ColorIndex = 3
End If

Next p

Cells(2, 15).Value = "Greatest % Incr"
Cells(3, 15).Value = "Greatest % Dec"
Cells(4, 15).Value = "Greatest Total Vol"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

For t = 2 To lastrow2

If Cells(t, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow2)) Then
    Cells(2, 16).Value = Cells(t, 10).Value
    Cells(2, 17).Value = Cells(t, 12).Value
    Cells(2, 17).NumberFormat = "0.00%"

ElseIf Cells(t, 12).Value = Application.WorksheetFunction.Min(Range("L2:L" & lastrow2)) Then
    Cells(3, 16).Value = Cells(t, 10).Value
    Cells(3, 17).Value = Cells(t, 12).Value
    Cells(3, 17).NumberFormat = "0.00%"
    
ElseIf Cells(t, 13).Value = Application.WorksheetFunction.Max(Range("M2:M" & lastrow2)) Then
    Cells(4, 16).Value = Cells(t, 10).Value
    Cells(4, 17).Value = Cells(t, 13).Value
    Cells(4, 17).NumberFormat = "0"
    
    End If

Next t

Next j


End Sub

