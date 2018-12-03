Sub FindFGTables2()

    Dim ws As Worksheet
    Dim rFound As Range
    Dim searchKeyword As String
    Dim inputWorkbookName As String
    Dim inputWorkbook As String
    Dim fnd As String, FirstFound As String
    Dim FoundCell As Range, rng As Range
    Dim myRange As Range, LastCell As Range
    Dim cel As Range, selectedRange As Range
    Dim ResultsDict As Dictionary


    Columns.EntireColumn.Hidden = False
    Rows.EntireRow.Hidden = False

    inputWorkbook = "Power and thermal COS Summary_ 20180821.xlsx"
    searchKeyword = "Red "

    If searchKeyword = "" Then Exit Sub

    'Cycle through each workbook
    'For Each ws In Workbooks(inputWorkbook).Worksheets
    For Each ws In Worksheets
        Debug.Print "WS: " & ws.Name
        With ws.UsedRange
            Set rng = Nothing
            Set myRange = Nothing
            Set LastCell = Nothing
            Set FoundCell = Nothing
            Set selectedRange = Nothing

            Set FoundCell = .Find(what:=searchKeyword, after:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart)

            'Check to see if anything was found
            If Not FoundCell Is Nothing Then
                FirstFound = FoundCell.Address
                'Debug.Print "Worksheet: " & ws.Name
                Debug.Print "", "ADDRESS: " & FirstFound
                Debug.Print "", "FOUND"

                'Check conditions for found
                'aka .FindNext(_)

            Else
                'GoTo NoMatch
                'Set rng = FoundCell
                Debug.Print "", "Not Found"
            End If

'Error Handler
NoMatch:
        'Debug.Print "No values were found in this worksheet: " & ws.Name
        End With

    Next ws

End Sub
