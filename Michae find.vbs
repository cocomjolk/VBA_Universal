Sub Table_Search_Power(pptSavePath, saveName, templateWbPath, pptTemplatePath)

    Dim templateWb As Workbook
    Dim pptApp As Object

    Dim inputWbPath As String
    Dim pptAppPath As String

    Dim ws As Worksheet
    Dim rFound As Range
    Dim searchKeyword As String
    Dim fnd As String, FirstFoundKeyWorkd As String
    Dim FoundCell As Range, rng As Range
    Dim myRange As Range, LastCell As Range
    Dim cel As Range, selectedRange As Range
    Dim resultsDict As Dictionary
    Dim tblDict As Dictionary
    Dim val1 As Byte
    Dim val2 As Byte
    Dim hiValue As Byte
    Dim inputWb As String
    Dim inputWsName As String

    Columns.EntireColumn.Hidden = False
    Rows.EntireRow.Hidden = False

    searchKeyword = "Red"
    If searchKeyword = "" Then Exit Sub
    'Cycle through all ws in each workbook
    For Each ws In Worksheets
        Debug.Print ws.name
        ws.Activate

        With ws.UsedRange

            'Cycle through results in tab
            'Used range of sheet to search'
            Set myRange = ActiveSheet.UsedRange
            'Cell count of used range
            Set LastCell = myRange.Cells(myRange.Cells.Count)
            Set FoundCell = myRange.Find(what:=searchKeyword, after:=LastCell, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)

            'Check to see if anything was found
            If Not FoundCell Is Nothing Then
                FirstFoundKeyWorkd = FoundCell.Address
                Set rng = FoundCell

                'Loop until cycled through all found keywords
                Do Until FoundCell Is Nothing
                    'Find next cell with fnd value
                    Set FoundCell = myRange.FindNext(after:=FoundCell)
                    'Add found cell to rng range variable
                    Set rng = Union(rng, FoundCell)
                    'Test to see if cycled through to first found cell
                    If FoundCell.Address = FirstFoundKeyWorkd Then Exit Do
                Loop

                'Select Cells Containing Find Value
                rng.Select

                Set selectedRange = Application.Selection

                For Each cel In selectedRange.Cells

                    tblName = cel.Offset(-1, 0).value
                    CellToRightVal = cel.Offset(0, 1).value
                    TwoCellsRightVal = cel.Offset(0, 2).value
                    isCellToRightNumeric = IsNumeric(CellToRightVal)
                    is2CellsToRightsNumeric = ((isCellToRightNumeric = False) And (IsNumeric(TwoCellsRightVal) = True))
                    CellDownVal = cel.Offset(1, 0).value
                    firstColumn = cel.Column
                    endColumnNum = cel.End(xlToRight).Column
                    tblRngLength = (endColumnNum - firstColumn)

                    'searches in string for "yellow"
                    isBelowAColor = (InStr(CStr(LCase(CellDownVal)), "yellow") > 0)

                    'determining where start column range begins
                    If ((is2CellsToRightsNumeric = True) And (isCellToRightNumeric = False)) Then
                        startColumn = ColumnNumberToLetter(cel.Column + 2)
                    Else
                        startColumn = ColumnNumberToLetter(cel.Column + 1)
                    End If

                    'finds end of column, converts from number to letter
                    endColumn = ColumnNumberToLetter(cel.End(xlToRight).Column)
                    'finds top row for range copy
                    startRow = cel.Row - 1
                    'finds bottom of range
                    endRow = cel.Row + 4

                    tblStatus = (((isCellToRightNumeric = True) Or (is2CellsToRightsNumeric = True) And (isCellToRightNumeric = False)) And isBelowAColor)
                    tblRng = startColumn & startRow & ":" & endColumn & endRow

                    Debug.Print "", "Table Status: " & tblStatus, ""

                    'If ((tblStatus = True) And (tblName <> "DSI summary") Or (tblName <> "dsi summary")) Then
                    If (tblStatus = True) Then
                        Debug.Print ""
                        Debug.Print "", "Keyword Value: " & cel.value
                        Debug.Print "", "Table Name: " & tblName
                        Debug.Print "", "Keyword Address: " & cel.Address
                        Debug.Print "", "Cell to right: " & CellToRightVal
                        Debug.Print "", "2 cells to right: " & TwoCellsRightVal
                        Debug.Print "", "Cell Down: " & CellDownVal
                        Debug.Print "", "Cell right Numeric: " & isCellToRightNumeric
                        Debug.Print "", "2 Cells right Numeric: " & is2CellsToRightsNumeric
                        Debug.Print "", "Cell down yellow?: " & isBelowAColor
                        Debug.Print "", "Table Range: " & tblRng

                        val1 = cel.Offset(0, 1).value
                        val2 = cel.Offset(0, 2).value
                        inputWb = ActiveWorkbook.name
                        inputWsName = ActiveSheet.name
                        Debug.Print "", "inputWb: " & inputWb
                        Debug.Print "", "inputWsName: " & inputWsName


                        If val1 >= val2 Then
                            hiValue = val1
                        Else
                            hiValue = val2
                        End If

                        Debug.Print "", "hiValue: " & hiValue

                        temp = Copy_Paste(tblRng, inputWb, inputWsName, tblName, templateWbPath, tblRngLength)
                        temp = CopyPaste_DSI_Tbl_To_PPT(pptTemplatePath, tblRngLength)
                        temp = Copy_Paste_Comment_Table(hiValue, tblName, templateWbPath, pptTemplatePath)

                    End If

                Next cel

                Debug.Print ""
            Else
                Debug.Print ""
                Debug.Print "", "No keyword matches"
                Debug.Print ""
            End If

        End With

    Next ws

    temp = PPT_SaveAs(pptSavePath, saveName)
    temp = Close_PPT_No_Save()
    Workbooks(inputWb).Close False
    temp = Quit_Application()


End Sub
