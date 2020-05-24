Attribute VB_Name = "Module1"
Sub stock_changes():

Dim total_change As Double
Dim percent_change As Double
Dim daily_change As Double
Dim change As Double
Dim total As Double

Dim days As Integer

Range("J1").Value = "Ticker"
Range("K1").Value = "Year Open"
Range("L1").Value = "Year Close"
Range("M1").Value = "Yearly Change"
Range("N1").Value = "Percent Change"
Range("O1").Value = "Total Stock Volume"
Range("J1").Font.Bold = True
Range("K1").Font.Bold = True
Range("L1").Font.Bold = True
Range("M1").Font.Bold = True
Range("N1").Font.Bold = True
Range("O1").Font.Bold = True

j = 0
total = 0
total_change = 0
percentage_change = 0
Start = 2

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To RowCount
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        j = j + 1
        total = Cells(i, 7)
        Range("J" & 2 + j).Value = Cells(i, 1).Value
        Range("K" & 2 + j).Value = Cells(i, 3).Value
        Range("L" & 2 + j).Value = 0
        Range("M" & 2 + j).Value = 0
        Range("N" & 2 + j).Value = "%" & 0
        Range("O" & 2 + j).Value = "$" & 0
    Else
            total = total + Cells(i, 7)
        End If
        
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Range("L" & 2 + j).Value = Cells(i, 6).Value()
        change = Range("L" & 2 + j).Value - Range("K" & 2 + j).Value
        Range("M" & 2 + j).Value = change
        percent_change = Round((change / Range("K" & 2 + j).Value), 2)
        Range("N" & 2 + j).Value = percent_change
        Range("O" & 2 + j).Value = total
        Select Case change
            Case Is > 0
                Range("M" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                Range("M" & 2 + j).Interior.ColorIndex = 3
            Case Else
                Range("M" & 2 + j).Interior.ColorIndex = 0
        End Select
        change = 0
        percentage_change = 0
        End If
        
    Next i
End Sub

