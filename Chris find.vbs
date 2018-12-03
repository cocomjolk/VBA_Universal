Sub SearchAllSheets()
    Dim ws As Worksheet
    Dim rFound As Range
    Dim searchKeyword As String
    Dim fnd As String, FirstFound As String
    Dim FoundCell As Range, rng As Range
    Dim myRange As Range, LastCell As Range
    Dim cel As Range, selectedRange As Range

    Columns.EntireColumn.Hidden = False
    Rows.EntireRow.Hidden = False

    searchKeyword = "Red ("
    If searchKeyword = "" Then Exit Sub
    'Cycle through each workbook
    For Each ws In Worksheets
        Debug.Print ws.Name
        ws.Activate
        With ws.UsedRange
            Set FoundCell = .Find(what:=searchKeyword, after:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart)

            'Cycle through results in tab
            Set myRange = ActiveSheet.UsedRange
            Set LastCell = myRange.Cells(myRange.Cells.Count)
            Set FoundCell = myRange.Find(what:=searchKeyword, after:=LastCell)

            'Check to see if anything was found
            If Not FoundCell Is Nothing Then
                FirstFound = FoundCell.Address
                Set rng = FoundCell

                'Loop until cycled through all unique finds
                Do Until FoundCell Is Nothing
                    'Find next cell with fnd value
                      Set FoundCell = myRange.FindNext(after:=FoundCell)

                    'Add found cell to rng range variable
                      Set rng = Union(rng, FoundCell)

                    'Test to see if cycled through to first found cell
                      If FoundCell.Address = FirstFound Then Exit Do
                Loop

            'Select Cells Containing Find Value
                rng.Select

                Set selectedRange = Application.Selection

                Dim colorArr(2) As String
                'these should be pulled from a config
                colorArr(0) = "yellow"
                colorArr(1) = "blue"
                colorArr(2) = "green"

                For Each cel In selectedRange.Cells

                    headCell = cel.Offset(-1, 0).value
                    nextColVal = cel.Offset(0, 1).value
                    isRightColumnAValue = IsNumeric(nextColVal)
                    belowRowVal = cel.Offset(1, 0).value
                    isBelowAColor = IsInArray(CStr(LCase(belowRowVal)), colorArr)

                    Debug.Print "", cel.Worksheet.Name, cel.value, cel.Address, "H: " & headCell, "R: " & nextColVal, "D: " & belowRowVal,
                    Debug.Print "isRight: " & isRightColumnAValue, "isBelowColor: " & isBelowAColor

                Next cel

            Else
                'GoTo NoMatch
                Debug.Print "", "No matches for: " & ws.Name
            End If


        End With
        'FoundCell = Empty
    Next ws
Exit Sub
'Error Handler
NoMatch:
  MsgBox "No values were found in this worksheet"

End Sub
