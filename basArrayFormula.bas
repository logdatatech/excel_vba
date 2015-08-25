
Sub array_formula()
    Dim rRange As Range, cell As Range
    Dim tot As Integer
    Set rRange = ActiveSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
        For Each cell In rRange
            If cell.HasArray Then
                MsgBox cell.Address & " " & cell.formula
                tot = tot + 1
             End If
         Next cell
    MsgBox "total number of array formula: " & tot
End Sub