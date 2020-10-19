Sub TagTxs()
'
' TagTxs Macro
'

    Dim myArray() As Variant
    Dim DataRange As Range
    Dim cell As Range
    Dim x As Long
    
    'Step1: Get the list of tags
    Sheets("Tags").Select
    
    Dim LastRowTagsSheet As Long
    With ActiveSheet
        LastRowTagsSheet = .Cells(.Rows.Count, "B").End(xlUp).Row
    End With
    
    'Determine the data you want stored
    Set DataRange = ActiveSheet.Range("B1:B" & LastRowTagsSheet)
    
    'Resize Array prior to loading data
    ReDim myArray(DataRange.Cells.Count)
    
    'Loop through each cell in Range and store value in Array
      For Each cell In DataRange.Cells
        myArray(x) = cell.Value
        x = x + 1
      Next cell
    
    'Print values to Immediate Window (Ctrl + G to view)
    Sheets("test").Select
    For x = LBound(myArray) To UBound(myArray) - 1
        Debug.Print myArray(x)
        Cells.Select
        If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
        Selection.AutoFilter
        ActiveSheet.Range("$A$1:$M$220").AutoFilter Field:=8, Criteria1:="*" & myArray(x) & "*", Operator:=xlFilterValues
        
        ' Get first and last visible filled cells.
        With Worksheets("test")
            With .Cells(1, 1).CurrentRegion
                'do all the .autofilter stuff here
                With .Resize(.Rows.Count - 1, .Columns.Count).Offset(1, 0)
                    If CBool(Application.Subtotal(103, .Cells)) Then
                        iFirstFilteredRow = _
                          .SpecialCells(xlCellTypeVisible).Rows(1).Cells.Row
                        '~~> rFirstFilteredRow is not a copy of the first visible row
                        'do something with rFirstFilteredRow
                    End If
                End With
            End With
        End With

        iLastFilteredRow = Cells(1, 1).CurrentRegion.SpecialCells(xlCellTypeVisible).End(xlDown).Row

        ' then do a filter using below formula
        With ActiveSheet
            Range("I" & iFirstFilteredRow & ":I" & iLastFilteredRow).Select
            ActiveCell.FormulaR1C1 = myArray(x)
            ' There is a bug with filldown that when first and last row are same, it gives wrong results.
            ' Thus skip autofill if there is only one row and code will work okay.
            If iLastFilteredRow > iFirstFilteredRow Then
                Selection.FillDown
            End If
        End With
        
    Next x

    ' Remove all filters now
    Cells.Select
    ActiveSheet.ShowAllData
    MsgBox "All txs with known tags, tagged"
End Sub
