Sub stockreport()

Dim WS_Count As Integer
Dim K As Integer
' Set WS_Count equal to the number of worksheets in the active
' workbook.
WS_Count = ActiveWorkbook.Worksheets.Count
MsgBox (WS_Count)
' Begin the loop.
For K = 1 To WS_Count
Worksheets(K).Activate


'add headers
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"

Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 15) = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Total Volume"

'add the list of stocks to column I
With ActiveSheet
    .Range("A2", .Range("A1").End(xlDown)).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=.Range("I2"), Unique:=True
    .Range("I2").Delete Shift:=xlShiftUp
End With

'Populate yearly change column

Dim lastprice As Double
Dim firstprice As Double
Dim yearchange As Double
Dim pctchange As Double


LastRowc1 = Cells(Rows.Count, 1).End(xlUp).Row
LastRowc9 = Cells(Rows.Count, 9).End(xlUp).Row

'find the initial first row of the first stock
firstprice = Cells(2, 6).Value
firstrow = 2
YrChangeRow = 2

    'iterate through first column to get first and last values
    For i = 2 To LastRowc1
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            lastprice = Cells(i, 6).Value
            lastrow = i
           
            'populate yearly change in column 10
            
            'This if clause was added to prevent a division by zero problem, since one of the stocks has all zeroes for all values
            If firstprice = 0 Then
                pctchange = firstprice
                Else: pctchange = ((lastprice - firstprice) / firstprice)
            End If
            Cells(YrChangeRow, 11).Value = pctchange
            
            yearchange = (lastprice - firstprice)
            Cells(YrChangeRow, 10).Value = yearchange
            
            totvol = Application.WorksheetFunction.Sum(Range(Cells(firstrow, 7), Cells(i, 7)))
            Cells(YrChangeRow, 12).Value = totvol
            
            If Cells(YrChangeRow, 10) > 0 Then
                Cells(YrChangeRow, 10).Interior.ColorIndex = 4
                Else: Cells(YrChangeRow, 10).Interior.ColorIndex = 3
            End If
            
            'reset firstprice variable for next stock
            YrChangeRow = YrChangeRow + 1
            firstprice = Cells(i + 1, 6).Value
            firstrow = i + 1
            
       End If
       
       
    Next i
        'Format Percentage column K
        Range(Cells(2, 11), Cells(LastRowc9, 11)).NumberFormat = "0.0%"
        
        'Populate the third table
        'find the last row in column K
        lr3 = Cells(Rows.Count, 11).End(xlUp).Row
        
        'find the greatest % increase and populate 3rd table
        
        maxstock = Application.WorksheetFunction.Max(Range(Cells(2, 11), Cells(lr3, 11)))
        Cells(2, 17) = maxstock
        
        'find the ticker for greatest % increase and populate 3rd table
        Row1 = Application.WorksheetFunction.Match(maxstock, Range("k1:k" & lr3), 0)
        Cells(2, 16) = Cells(Row1, 9)
        
        'find the greatest % decrease and populate 3rd table
        minstock = Application.WorksheetFunction.Min(Range(Cells(2, 11), Cells(lr3, 11)))
        Cells(3, 17) = minstock
        
        'find the ticker for greatest % decrease  and populate 3rd table
        Row1 = Application.WorksheetFunction.Match(minstock, Range("k1:k" & lr3), 0)
        Cells(3, 16) = Cells(Row1, 9)
        
        
        'find the greatest volume and populate 3rd table
        maxvol = Application.WorksheetFunction.Max(Range(Cells(2, 12), Cells(lr3, 12)))
        Cells(4, 17) = maxvol
        
        'find the ticker for greatest % decrease  and populate 3rd table
        Row1 = Application.WorksheetFunction.Match(maxvol, Range("L1:l" & lr3), 0)
        Cells(4, 16) = Cells(Row1, 9)
        
        'format the two percentages
        Range("q2:q3").NumberFormat = "0.0%"
        
        'Autofit columns for appearance
        Columns("A:Q").Select
        Range("Q1").Activate
        Selection.Columns.AutoFit
        
   Next K

End Sub

