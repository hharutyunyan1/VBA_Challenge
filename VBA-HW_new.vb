Sub Stock_Checker()

Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
WS.Activate


Dim ticker As String
Dim yearly_change As Double
Dim opening_price As Double
Dim closing_price As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim row As Long
Dim column As Long

Cells(1, "I").Font.Bold = True
Cells(1, "J").Font.Bold = True
Cells(1, "K").Font.Bold = True
Cells(1, "L").Font.Bold = True
Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Yearly Change"
Cells(1, "K").Value = "Percent Charge"
Cells(1, "L").Value = "Total Stock Volume"

row = 2
column = 1
summaryRowCount = 2
total_stock_volume = 0

opening_price = Cells(2, column + 2).Value

For row = 2 To Range("A2").End(xlDown).row

    If Cells(row, 1).Value <> Cells(row + 1, 1).Value Then
         closing_price = Cells(row, 6).Value
         yearly_change = closing_price - opening_price
         Cells(summaryRowCount, column + 9).Value = yearly_change
            If (opening_price = 0 And closing_price = 0) Then
                percent_change = 0
            ElseIf (opening_price = 0 And closing_price <> 0) Then
                percent_change = 1
            Else
                percent_change = yearly_change / opening_price
                Cells(summaryRowCount, 11).Value = percent_change
                Cells(summaryRowCount, 11).NumberFormat = "0.00%"
            End If
         opening_price = Cells(row + 1, column + 2).Value
    End If

    If Cells(row, 1).Value <> Cells(row + 1, 1).Value Then
        total_stock_volume = total_stock_volume + Cells(row, 7).Value
        Cells(summaryRowCount, 9).Value = Cells(row, 1).Value
        Cells(summaryRowCount, 12).Value = total_stock_volume
        summaryRowCount = summaryRowCount + 1
        total_stock_volume = 0
    Else
        total_stock_volume = total_stock_volume + Cells(row, 7).Value
    End If
Next row

For row = 2 To Range("J2").End(xlDown).row

    If Cells(row, 10).Value > 0 Then
        Cells(row, 10).Interior.ColorIndex = 4
    Else
        Cells(row, 10).Interior.ColorIndex = 3
    End If
    
Cells(2, "O").Font.Bold = True
Cells(3, "O").Font.Bold = True
Cells(4, "O").Font.Bold = True
Cells(1, "P").Font.Bold = True
Cells(1, "L").Font.Bold = True
Cells(2, "O").Value = "Greatest % Increase"
Cells(3, "O").Value = "Greatest % Decrease"
Cells(4, "O").Value = "Greatest Total Volume"
Cells(1, "P").Value = "Ticker"
Cells(1, "Q").Value = "Value"

'Finding Greatest % (Increase/Decrease) AND Greatest Total Volume
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=MAX(RC[-6]:R[705712]C[-6])"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "=MIN(R[-1]C[-6]:R[705711]C[-6])"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=MAX(R[-2]C[-5]:R[705710]C[-5])"
    Range("Q5").Select
    Cells.Find(What:="-95.73", After:=ActiveCell, LookIn:=xlValues, LookAt _
        :=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    ActiveWindow.ScrollRow = 646
    ActiveWindow.ScrollRow = 1
    Range("P3").Select
    ActiveCell.FormulaR1C1 = "CBO"
    Range("Q2").Select
    ActiveWindow.ScrollRow = 646
    ActiveWindow.ScrollRow = 1
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "DM"
    Range("Q4").Select
    ActiveWindow.ScrollRow = 1
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "BAC"
    Range("O9").Select

Next row
Next WS
End Sub