'SN1476 Modules
Public logPath As String
Public xmlPath As String
Public PartList As Dictionary
Public Parts() As String
Public ExceptionList As Dictionary

Set logPath = "C:\Users\michael_quiroz\Desktop\SN1479\Logs\"
Set xmlPath = "C:\Users\michael_quiroz\Desktop\SN1479\Config\"

Sub setXmlPath(Path)
    xmlPath = Path
End Sub

Sub setLogPath(Path)
    logPath = Path
    Debug.Print logPath
End Sub



Sub Parent()

    Dim wsGDS As String
    Dim wbGDS As String

    wsGDS = "Weekly GDS"
    wbGDS = "Book1"

'formats and adds exemption tab
PrepGDS

'Capture all PNs from Tools_Output_Suppliers
CapturePNData


'Paste To Dashboard
TranferPartData


'Paste Exceptions
HandleExceptions


'Cleanup

End Sub

Sub PrepGDS()

    'need to get file name
    'Open GDS and DSI summary file
    'create tab in GDS file expectations'

    Dim wsGDS As String
    Dim wbGDS As String

    wsGDS = "Weekly GDS"
    wbGDS = "Book1"

    Call Activate_GDS_Sheet(wbGDS, wsGDS)

    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AN$1").AutoFilter field:=3, Criteria1:="WW"
    ActiveSheet.Range("$A$1:$AN$1").AutoFilter field:=7, Criteria1:= _
    "Projected DSI"

    ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
    ActiveSheet.name = "Exceptions"
    Call Activate_GDS_Sheet(wbGDS, wsGDS)

End Sub

Sub CapturePNData()
    Dim currentBook As Workbook
    Dim currentSheet As Worksheet
    Dim ColNum As Integer
    Dim ColRow As Integer
    Dim rngAf As Range
    Dim CurrentDict As Dictionary
    'Perform filtering and sorting

    'Find Range only visible
        For Each rngRow In Intersect(ActiveSheet.UsedRange, ActiveSheet.Range("AC2:" & LastRow).SpecialCells(xlCellTypeVisible))
        currentRow = rngRow.Row
        Set CurrentDict = New Dictionary

        If currentRow <> 1 Then
          'map to vars
          partNum = Range(ColumnNumberToLetter(partNumCol) & currentRow).Value
          partDesc = Range(ColumnNumberToLetter(partDescCol) & currentRow).Value
          Platform = Range(ColumnNumberToLetter(platformCol) & currentRow).Value
          demandRegion = Range(ColumnNumberToLetter(demandRegionCol) & currentRow).Value
          siteFactory = Range(ColumnNumberToLetter(siteFactoryCol) & currentRow).Value

          Commodity = partDesc


          CurrentDict.Add "descr", partDesc
          CurrentDict.Add "platform", Platform
          CurrentDict.Add "commodity", Commodity
          CurrentDict.Add "demand_region", demandRegion
          CurrentDict.Add "site", siteFactory

          If Not PartsDict.Exists(partNum) Then
            PartsDict.Add partNum, CurrentDict
          End If

          Debug.Print "Part Info..."
          WriteToLog ("Part Info...")
        End If
    Next rngRow



    'Store Data

    'Example of using a dictionary
    'PartList(partNum).Add "row", "row #"
    'PartList(partNum).Item("row") = Platform

End Sub

Sub TranferPartData()

    'Find Part

    'If Part Not Found Add to Exception

    'ExceptionList(partNum).Add "data", "whatever data"


    'Paste Data


End Sub

Sub HandleExceptions()

    'Create or update Exception tab

    'Loop through parts and fill out sheet

    For Each exceptionPart In ExceptionList.Keys()
        partRow = ExceptionList.Item(exceptionPart).Item("row")

    Next

End Sub


Function Save_GDS_Name()

  GDSwbName = ActiveWorkbook.Name
  GDSwsName = ActiveSheet.Name

  Debug.Print "WB Name:", GDSwbName
  Debug.Print "WS Name:", GDSwsName
  ActiveWorkbook.SaveAs ("Z:\SN1213 PPT Creation\GDS\WeeklyGDS")
  ActiveWorkbook.Close False

End Function


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
        If listNode Is Nothing Then
            Debug.Print "Nothing found for xPath: ", nodePath
            SingleNodeText = ""
        Else
            SingleNodeText = listNode.Text
        End If

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

Function WriteToLog(strText)

    Dim FSO As New FileSystemObject
    Dim ts As TextStream
    Dim strLogPath As String
    Dim strFileName As String

    '- check if file exists
    If FSO.FileExists(logPath) Then

        '- open this file to write to it
        Set ts = FSO.OpenTextFile(logPath, ForAppending)
        ts.WriteBlankLines 0
        ts.WriteLine "[" & Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "] " & CStr(strText)

    Else
        '- create file and write to it - can set new file name here!
        Set ts = FSO.CreateTextFile(logPath, True)
        ts.WriteLine "[" & Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "] " & "LogFile Creates"
        ts.WriteLine "[" & Strings.Format(Now, "YYYY-MM-DD HH:nn:ss") & "] " & CStr(strText)

    End If

    '- close down the file
    ts.Close
End Function

Function IsWbOpen(wbName As String)
On Error Resume Next
IsWbOpen = (Len(Workbooks(wbName).Name) > 0)
End Function

Sub RemoveFilters(ByRef WhichSheet As Worksheet)

    If WhichSheet.FilterMode Then WhichSheet.ShowAllData
    If WhichSheet.AutoFilterMode Then WhichSheet.AutoFilterMode = False

End Sub

Function NewestFile(Directory As Variant, Optional FileSpec = "*.*")

Dim FileName As String
Dim MostRecentFile As String
Dim MostRecentDate As Date

'specify the directory
FileName = Dir(Directory & FileSpec)

If FileName <> "" Then
    MostRecentFile = FileName
    MostRecentDate = FileDateTime(Directory & FileName)
    Do While FileName <> ""
        If FileDateTime(Directory & FileName) > MostRecentDate Then
             MostRecentFile = FileName
             MostRecentDate = FileDateTime(Directory & FileName)
             Debug.Print MostRecentFile
        End If
        FileName = Dir
    Loop
End If

NewestFile = MostRecentFile

End Function

Function GetLatestFolder(Path As String)
    Dim FSO, FS, F, DtLast As Date, Result As String
    Set FSO = CreateObject("scripting.FileSystemObject")
    Set FS = FSO.GetFolder(Path).SubFolders
    For Each F In FS
        If F.DateCreated > DtLast Then
             DtLast = F.DateCreated
             Result = F.Name
        End If
    Next
    GetLatestFolder = Result
End Function

Function Activate_GDS_WB(wbName)
  Windows(wbName).Activate
End Function

Function Activate_DSI_Summary_WB(wbName)
  Windows(wbName).Activate
End Function

Function Activate_GDS_Sheet(wbName, wsName)
  Windows(wbName).Activate
  Sheets(wsName).Select
End Function

Function Activate_DSI_Original_Sheet(wbName, wsName)
  Windows(wbName).Activate
  Sheets(wsName).Select
End Function
