Public log As TextStream
Public logPath As String
Public xmlPath As String

Sub setXmlPath(path)
    xmlPath = path
End Sub

Sub setLogPath(path)
    logPath = path
    Debug.Print logPath
End Sub

Sub getXMLPath()
    Debug.Print xmlPath
    WriteToLog xmlPath
End Sub

'@Description: Loop through workbook and all tabs to search for specified keywords and return array of founds ranges
'@Assumption: Excel file is open and active
Sub FindFGTables()
    Dim ws As Worksheet
    Dim rFound As Range
    Dim searchKeyword As String
    Dim fnd As String, FirstFound As String
    Dim FoundCell As Range, rng As Range
    Dim myRange As Range, LastCell As Range
    Dim cel As Range, selectedRange As Range

    Columns.EntireColumn.Hidden = False
    Rows.EntireRow.Hidden = False

    searchKeyword = "Red"
    'TBD: change from static keyword to using config to pull search words
    'TBD: loop through all keywords
    'delimiterRecord = GetXMLData(xmlPath, "//Config/delimiters/record")

    If searchKeyword = "" Then Exit Sub
    'Cycle through each workbook
    For Each ws In Worksheets
        Debug.Print ws.Name
        With ws.UsedRange
            Set FoundCell = .Find(what:=searchKeyword, after:=.Cells(1, 1), LookIn:=xlValues, LookAt:=xlWhole)

            'Cycle through results in tab
            Set myRange = ActiveSheet.UsedRange
            Set LastCell = myRange.Cells(myRange.Cells.Count)
            Set FoundCell = myRange.Find(what:=searchKeyword, after:=LastCell)

            'Check to see if anything was found
            If Not FoundCell Is Nothing Then
                FirstFound = FoundCell.Address
            Else
                GoTo NoMatch
            End If

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

            For Each cel In selectedRange.Cells

            headCell = cel.Offset(-1, 0).Value

            Debug.Print "", cel.Value, cel.Address, headCell

            Next cel


        End With
    Next ws

Exit Sub

'Error Handler
NoMatch:
  MsgBox "No values were found in this worksheet"

End Sub


'@input("A")
'@return int 1
Function ColumnLetterToNumber(ColumnLetter)
    ColumnLetterToNumber = Range(ColumnLetter & 1).Column
End Function

'@input("1")
'@return str "a"
Function ColumnNumberToLetter(ColNum)
    ColumnNumberToLetter = Split(Cells(1, ColNum).Address, "$")(1)
End Function

Sub CopyToClipboard(val)
    Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    With objData
      .SetText val
      .PutInClipboard
    End With
End Sub

'xmlPath, nodePath
Function GetXMLData(xmlPath, nodePath, Optional attribName = "")
    Dim XDoc As Object
    Dim lists As Object
    Dim listNode  As Object
    Dim fieldNode As Object

    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (xmlPath)

    If XDoc.Load(xmlPath) Then
        'Get Document Elements
        Set lists = XDoc.DocumentElement
        Set listNode = lists.SelectSingleNode(nodePath)
        SingleNodeText = listNode.Text
        'MsgBox SingleNodeText
    Else
        GoTo NoLoad
    End If

    Set XDoc = Nothing

    GetXMLData = SingleNodeText
Exit Function
'Error Handler
NoLoad:
  Dim xPE        As Object
  Dim strErrText As String
  Set xPE = XDoc.parseError
  With xPE
     strErrText = "Load error " & .ErrorCode & " xml file " & vbCrLf & _
     Replace(.URL, "file:///", "") & vbCrLf & vbCrLf & _
    xPE.reason & _
    "Source Text: " & .srcText & vbCrLf & vbCrLf & _
    "Line No.:    " & .Line & vbCrLf & _
    "Line Pos.: " & .linepos & vbCrLf & _
    "File Pos.:  " & .filepos & vbCrLf & vbCrLf
  End With
  MsgBox strErrText, vbExclamation
  Set xPE = Nothing
  Exit Function
End Function

Function UpdateXML(xmlPath, nodePath, nodeValue)
    Dim XDoc As Object
    Dim lists As Object
    Dim listNode  As Object
    Dim fieldNode As Object

    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load (xmlPath)

    If XDoc.Load(xmlPath) Then
        'Get Document Elements
        Set lists = XDoc.DocumentElement

        Set listNode = XDoc.SelectSingleNode(nodePath)
        listNode.Text = nodeValue
        XDoc.Save (xmlPath)
    Else
        GoTo NoLoad
    End If

    Set XDoc = Nothing

Exit Function
'Error Handler
NoLoad:
  Dim xPE        As Object
  Dim strErrText As String
  Set xPE = XDoc.parseError
  With xPE
     strErrText = "Load error " & .ErrorCode & " xml file " & vbCrLf & _
     Replace(.URL, "file:///", "") & vbCrLf & vbCrLf & _
    xPE.reason & _
    "Source Text: " & .srcText & vbCrLf & vbCrLf & _
    "Line No.:    " & .Line & vbCrLf & _
    "Line Pos.: " & .linepos & vbCrLf & _
    "File Pos.:  " & .filepos & vbCrLf & vbCrLf
  End With
  MsgBox strErrText, vbExclamation
  WriteToLog "ERROR: " & strErrText
  Set xPE = Nothing
  Exit Function
End Function

Function cleanFileName(newFileName)
    newFileName = Replace(newFileName, "|", "_")
    newFileName = Replace(newFileName, "/", "_")
    newFileName = Replace(newFileName, "\", "_")
    newFileName = Replace(newFileName, "%", "")
    newFileName = Replace(newFileName, "*", "")
    newFileName = Replace(newFileName, ":", "_")
    newFileName = Replace(newFileName, ">", "")
    newFileName = Replace(newFileName, "<", "")
    newFileName = Replace(newFileName, ".", "_")

    cleanFileName = newFileName
End Function


Function CopyRange(rngStr)
    ActiveSheet.Range(rngStr).Select
    Selection.Copy
End Function

Function CopyRangeToPicture(rngStr)
    ActiveWorkbook.ActiveSheet.Range(rngStr).CopyPicture Appearance:=xlScreen, Format:=xlPicture
End Function

Function PasteToRange(rngStr)
    'TBD
End Function

Function WriteToLog(strText)

    Dim fso As New FileSystemObject
    Dim ts As TextStream
    Dim strLogPath As String
    Dim strFileName As String

    '- check if file exists
    If fso.FileExists(logPath) Then

        '- open this file to write to it
        Set ts = fso.OpenTextFile(logPath, ForAppending)
        ts.WriteBlankLines 0
        ts.WriteLine "[" & Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "] " & CStr(strText)

    Else
        '- create file and write to it - can set new file name here!
        Set ts = fso.CreateTextFile(logPath, True)
        ts.WriteLine "[" & Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "] " & "LogFile Creates"
        ts.WriteLine "[" & Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "] " & CStr(strText)

    End If

    '- close down the file
    ts.Close
End Function
