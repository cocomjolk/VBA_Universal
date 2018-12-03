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

'@Description Find all at risk parts, define at risk type and output info
Sub CodeRedAlg()
Dim fnd As String, FirstFound As String
Dim FoundCell As Range, rng As Range
Dim myRange As Range, LastCell As Range
Dim cel As Range, selectedRange As Range
Dim dsiAdd As Variant, dsiRow As Integer, headerRow As Integer
Dim partName As String, partDescr As String, outlookStatus As String
Dim isLowDsi As Boolean, isPSTMRP As Boolean, isSTBL As Boolean
Dim CurrentWk As Variant, CurrentWkVal As Variant, PreviousWkVal As Variant, FirstWk As Variant
Dim lowDsiMin As Integer, lowDsiMax As Integer
Dim delimiterRecord As String, delimiterValue As String
Dim alertStr As String

Dim wkCounter As Integer

'Pull Data From Config
delimiterRecord = GetXMLData(xmlPath, "//Config/delimiters/record")
delimiterValue = GetXMLData(xmlPath, "//Config/delimiters/value")
lowDsiMin = GetXMLData(xmlPath, "//Config/codeRedAlg/low_dsi/min")
lowDsiMax = GetXMLData(xmlPath, "//Config/codeRedAlg/low_dsi/max")
Update = UpdateXML(xmlPath, "//Config/codeRedAlg/cache/data", "")

alertStr = Empty

Columns.EntireColumn.Hidden = False
Rows.EntireRow.Hidden = False

Set myRange = ActiveSheet.UsedRange
Set LastCell = myRange.Cells(myRange.Cells.Count)
Set FoundCell = myRange.Find(what:="Project DSI", after:=LastCell)

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
        dsiAdd = cel.Address
        dsiRow = cel.Row
        partName = cel.Offset(0, -2).Value
        partDescr = cel.End(xlUp).Offset(1, -2).Value
        headerRow = cel.End(xlUp).Row
        headerValue = cel.End(xlUp).Value
        wkCounter = 1

        Debug.Print partName, dsiAdd, cel.Value, dsiRow, headerRow, headerValue

        'Define outlook status
        isSTBL = Empty
        isPSTMRP = Empty
        isLowDsi = Empty

        ''IF MATCHES LOW DSI
        For Each wkCel In Range("E" & dsiRow & ":Q" & dsiRow).Cells

            currentCol = Split(wkCel.Address, "$")(1)
            CurrentWk = Range(currentCol & headerRow).Value

            'If there's a cell error
            If (TypeName(wkCel.Value) = "Error") Then
                currentVal = 0
            Else
                currentVal = Int(wkCel.Value)
            End If

            If (wkCounter = 1) Then
                FirstWk = currentVal
                CurrentWkVal = currentVal

                'Check if first week is negative
                If (currentVal < 0) Then
                    isSTBL = True
                End If

            Else
                'Check week data
                PreviousWkVal = CurrentWkVal
                CurrentWkVal = currentVal

                'Debug.Print "WK: " & wkCounter & "--", lowDsiMin & "<" & PreviousWkVal & " And " & PreviousWkVal & "<" & lowDsiMax, lowDsiMin & "<" & CurrentWkVal & " And " & CurrentWkVal & "<" & lowDsiMax

                If (currentVal < 0) Then
                    isPSTMRP = True
                End If


                If ((lowDsiMin < PreviousWkVal And PreviousWkVal < lowDsiMax) And (lowDsiMin < CurrentWkVal And CurrentWkVal < lowDsiMax)) Then
                    isLowDsi = True
                End If

            End If

            wkCounter = wkCounter + 1
            'Debug.Print "", wkCel.Address, CurrentWk, "Low Value: " & currentVal
        Next wkCel

        ''LABEL DSI STATUS
            outlookStatus = Empty

            If (isSTBL = True) Then
                outlookStatus = "STBL"
            ElseIf (isPSTMRP = True And (outlookStatus = Empty)) Then
                outlookStatus = "Potential Short to MRP"
            ElseIf (isLowDsi = True And (outlookStatus = Empty)) Then
                outlookStatus = "Low DSI"
            Else
                outlookStatus = "OK"
            End If


        ''CONVERT status info to string
        If outlookStatus <> "OK" Then

            alertValue = dsiAdd & delimiterValue & partName & delimiterValue & partDescr & delimiterValue & outlookStatus & delimiterValue & headerRow
            If alertStr = Empty Then
                alertStr = alertValue
            Else
                alertStr = alertStr & delimiterRecord & alertValue
            End If

        End If


        'Debug.Print "", partName, dsiAdd, cel.Value, dsiRow, "Status: " & outlookStatus
        WriteToLog "    " & partName & "  " & partDescr & "  " & "Status: " & outlookStatus & "  " & "[LDSI: " & isLowDsi & "|PSTMRP:" & isPSTMRP & "|STBL:" & isSTBL & "]"
        Debug.Print "", partName, partDescr, "Status: " & outlookStatus, "[LDSI: " & isLowDsi & "|PSTMRP:" & isPSTMRP & "|STBL:" & isSTBL & "]"
        'Debug.Print partName, outlookStatus
        Debug.Print ""
    Next cel

        Debug.Print "-------------------"
        Debug.Print "Code Red Alert Data:"
        Debug.Print alertStr

        Update = UpdateXML(xmlPath, "//Config/codeRedAlg/cache/data", alertStr)

Exit Sub

'Error Handler
NoMatch:
  MsgBox "No values were found in this worksheet"
End Sub

Sub CreateAllFiles()
    Dim alertList As Dictionary
    Dim partDict As Dictionary
    Dim CalDict As Dictionary
    Dim FilePaths As String

    Dim partNumStr As String
    Set alertList = AlertStrToDict()
    Set CalDict = getCalendarDict()

    delimiterRecord = GetXMLData(xmlPath, "//Config/delimiters/record")
    delimiterValue = GetXMLData(xmlPath, "//Config/delimiters/value")
    Update = UpdateXML(xmlPath, "//Config/codeRedAlg/cache/templateFiles", "")

    Filename = ActiveWorkbook.Name
    Platform = Trim(Split(Filename, "GDS")(0))

    For Each partNum In alertList.Keys()
        partNumStr = partNum
        newFilePath = CreateCodeRedFile(alertList(partNum), partNumStr)
        alertList(partNum).Add "filePath", newFilePath
        alertList(partNum).Item("Platform") = Platform
        tmp = UpdateCodeRedFile(alertList(partNum), CalDict, partNumStr)

        If FilePaths = Empty Then
            FilePaths = newFilePath
        Else
            FilePaths = FilePaths & delimiterRecord & newFilePath
        End If
    Next

    Update = UpdateXML(xmlPath, "//Config/codeRedAlg/cache/templateFiles", FilePaths)

End Sub

Function CreateCodeRedFile(partDict As Dictionary, partNumber As String)
    Dim templatePath As String, savePath As String
    Dim Platform As String, Commodity As String
    Dim CurrentDate As String

    'Pull Data From Config
    templatePath = GetXMLData(xmlPath, "//Config/codeRedAlg/templateFilePath")
    savePath = GetXMLData(xmlPath, "//Config/codeRedAlg/savePath")

    partNumber = partNumber
    Commodity = partDict.Item("Commodity")

    CommSplit = Split(Commodity, ",")
    CurrentDate = Format(DateTime.Now, "yyyy_MM_dd")

    'Copy templateFile / Rename to file format (PN_commodity_date)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    ' First parameter: original location\file
    ' Second parameter: new location\file
    CommSize = UBound(CommSplit) - LBound(CommSplit) + 1

    If (CommSize > 1) Then
        newFileName = partNumber & "_" & CommSplit(0) & CommSplit(1) & "_" & CurrentDate
    Else
        newFileName = partNumber & "_" & CommSplit(0) & "_" & CurrentDate
    End If
    'replace any special characters like & / \
    newFileName = cleanFileName(newFileName)

    newFilePath = savePath & newFileName & ".xlsx"
    objFSO.CopyFile templatePath, newFilePath

    CreateCodeRedFile = newFilePath
End Function

Function UpdateCodeRedFile(partDict As Dictionary, CalDict As Dictionary, partNumber As String)
    Dim gdsBook As Excel.Workbook
    Dim codeRedBook As Excel.Workbook
    Dim codeRedSheet As Excel.Sheets
    Dim TemplateMapping As Dictionary
    Dim filePath As Variant, WarpFilePath As Variant
    Dim TemplateTabName As Variant
    Dim posPn As String, posLob As String

    filePath = partDict.Item("filePath")
    Set TemplateMapping = getTemplateMapping()
    TemplateTabName = TemplateMapping.Item("tab_name")

    'Insert Common Data
    partNumber = partNumber
    Update = UpdateXML(xmlPath, "//Config/codeRedAlg/cache/part", partNumber)
    Update = UpdateXML(xmlPath, "//Config/codeRedAlg/cache/currentFile", filePath)

    Debug.Print partNumber, filePath, TemplateTabName

    Set gdsBook = Application.ActiveWorkbook
    Set codeRedBook = Workbooks.Open(filePath)

    codeRedBook.Sheets(TemplateTabName).Activate

    Range(TemplateMapping.Item("pn")).Value = partNumber
    Range(TemplateMapping.Item("commodity")).Value = partDict.Item("Commodity")
    Range(TemplateMapping.Item("platform")).Value = partDict.Item("Platform")
    Range(TemplateMapping.Item("alert_level")).Value = partDict.Item("Status")

    'DSI Recovery
    headerRow = partDict("Header")
    dsiRow = Split(partDict("Position"), "$")(2)
    gdsBook.Sheets("DS_Outlook_Report").Activate
    dsiRecover = GetRecoveryDSI(Int(headerRow), Int(dsiRow))
    codeRedBook.Sheets(TemplateTabName).Activate
    Range(TemplateMapping.Item("recover_from_dsi")).Value = dsiRecover

    'Read GDS file to capture screenshot
    gdsBook.Sheets("DS_Outlook_Report").Activate
    cpyRng = "A" & partDict.Item("Header") & ":Q" & dsiRow
    'Debug.Print "Copy Range: " & cpyRng

    Range(cpyRng).Columns.AutoFit
    Range(cpyRng).CopyPicture xlScreen, xlPicture
    codeRedBook.Sheets(TemplateTabName).Activate
    codeRedBook.Sheets(TemplateTabName).Paste _
    Destination:=Worksheets(TemplateTabName).Range(TemplateMapping.Item("gds"))

    'Update Current Supply
    gdsBook.Sheets("DS_Outlook_Report").Activate
    CurrentSupply = GetCurrentSupply(partDict, CalDict, partNumber)
    Debug.Print "Current SPLY Val: ", CurrentSupply
    codeRedBook.Sheets(TemplateTabName).Activate
    Range(TemplateMapping.Item("current_supply")).Value = CurrentSupply

    'Waterfall screenshot


    'Close Excel file
    codeRedBook.Close True
End Function

Function GetCurrentSupply(partDict As Dictionary, CalDict As Dictionary, partNumber As String, Optional partType As String = "NB")
    Dim gdsBook As Excel.Workbook
    Dim FindRng As Range, TotalRng As Range, LastCell As Range, wkCel As Range
    Dim WkCount As Integer, MonthCount As Integer
    Dim monthWks As Integer

    Set gdsBook = Application.ActiveWorkbook
    gdsBook.Sheets("DS_Outlook_Report").Activate
    Position = partDict.Item("Position")
    Header = partDict.Item("Header")

    Set LastCell = Range("C" & Header)
    Set FindRng = Range("C" & Header & ":" & Position)
    Debug.Print "Current Supply INF: ", "C" & Header & ":" & Position

    Set TotalRng = FindRng.Find(what:="Supply Total", after:=LastCell)

    If TotalRng Is Nothing Then
        'error
    End If

    TotalRow = TotalRng.Row
    StartRow = Range("E" & Header).Row
    MonthCount = 0
    CurrentMonth = Empty
    previousMonth = Empty
    outputStr = Empty

    Debug.Print "CSF Rows: ", "Total Row: " & TotalRow, "Start Row: " & StartRow

    'Start looping through each week header to capture the months
    For Each wkCel In Range("E" & dsiRow & ":Q" & dsiRow).Cells
        'is month different
        wkDate = wkCel.Value
        currentColumn = ColumnNumberToLetter(wkCel.Column)
        wkValue = Range(currentColumn & TotalRow).Value
        CurrentMonth = Left(wkDate, 1)

        wkValue = WorksheetFunction.RoundUp((Int(wkValue) / 1000), 0) & "K"
        TranslatedDate = CalDict(wkDate).Item("fiscalMW")

        If ((previousMonth <> CurrentMonth) And (previousMonth <> Empty)) Then
            MonthCount = MonthCount + 1
        End If

        If (partType = "NB" And MonthCount = 1) Then
            Exit For
        ElseIf (MonthCount = 2) Then
            Exit For
        End If

        If outputStr = Empty Then
            outputStr = wkValue & " in " & TranslatedDate
        Else
            outputStr = outputStr & ", " & wkValue & " in " & TranslatedDate
        End If

        Debug.Print "", "GetCurrentSupply: " & MonthCount, previousMonth, CurrentMonth, "" & currentColumn & TotalRow
        previousMonth = CurrentMonth

    Next wkCel

    GetCurrentSupply = outputStr

End Function

Function GetRecoveryDSI(headerRow As Integer, dsiRow As Integer)
    Dim wkCel As Range
    Dim redAlert As Boolean
    Dim recovered As Boolean
    Dim redAlertPosition As Variant
    Dim currentColor As Variant, previousColor As Variant
    Dim wkCounter As Integer, alertCounter As Integer
    Dim red As Variant, green As Variant, blue As Variant

    green = 65280
    red = 255
    blue = 16737843

    wkCounter = 1
    alertCounter = 0
    currentColor = Empty
    previousColor = Empty
    redAlert = Empty
    redAlertPosition = Empty
    CurrentWk = Empty

    For Each wkCel In Range("E" & dsiRow & ":Q" & dsiRow).Cells

                currentCol = Split(wkCel.Address, "$")(1)
                CurrentWk = Range(currentCol & headerRow).Value

                If (wkCounter <> 1) Then
                    previousColor = currentColor
                End If

                currentColor = wkCel.DisplayFormat.Interior.Color

                If (currentColor = red) Then
                    redAlert = True
                    alertCounter = alertCounter + 1
                    redAlertPosition = wkCel.Address
                    redAlertColumn = ColumnNumberToLetter(wkCel.Column)
                End If

                'Debug.Print CurrentWk, "Previous: " & previousColor, "Current: " & currentColor, wkCel.Address, wkCel.Value

                wkCounter = wkCounter + 1
    Next

    Debug.Print redAlert, alertCounter, redAlertPosition

    If redAlert = True Then
        CurrentWk = Empty
        recovered = Empty

        For Each wkCel In Range(redAlertColumn & dsiRow & ":Q" & dsiRow).Cells
            currentColor = wkCel.DisplayFormat.Interior.Color

            If ((currentColor = blue Or currentColor = green) And recovered = Empty) Then
                    currentCol = Split(wkCel.Address, "$")(1)
                    CurrentWk = Range(currentCol & headerRow).Value
                    recovered = True
            End If
        Next

        'Cycle through weeks after the most distant red cell and check if ever recovered
        If recovered = True And CurrentWk <> Empty Then
            Debug.Print "Recovered on: " & CurrentWk
            GetRecoveryDSI = CurrentWk
        Else
            Debug.Print "Recovery TBD"
            GetRecoveryDSI = "TBD"
        End If

    Else
        Debug.Print "Recovery TBD"
        GetRecoveryDSI = "TBD"
    End If

End Function

Function initWarpShot()
    'Setup warp to bypass creds prompt and update prompt

End Function


Function captureWarpShot()
    Dim gdsBook As Excel.Workbook
    Dim warpBook As Excel.Workbook
    Dim Calendar As Worksheet
    Dim Review As Worksheet
    Dim cel As Range
    Dim yr As String
    Dim q As String
    Dim w As String
    Dim Version As String

    partNum = "DMYJX"
    WarpFilePath = getWarpFilePath()

    Set gdsBook = Application.ActiveWorkbook

    Set warpBook = Workbooks.Open(WarpFilePath)

    'Enter PN
    warpBook.Sheets("Detailed Review").Activate
    Application.Calculate
    Range("A35").Value = partNum

    Set Calendar = Sheets("Calendar")
    Set Review = Sheets("Detailed Review")

    Version = Calendar.Cells(8, 2)
    yr = Left(Version, 4)
    q = Mid(Version, 5, 2)
    w = Mid(Version, 7, Len(Version) - 6)

    Review.Cells(3, 5) = yr
    Review.Cells(3, 6) = q
    Review.Cells(3, 7) = w

    'Press Go
    callMethod = "'" & warpBook.Name & "'!HyperionUpdate"
    Debug.Print callMethod
    Application.DisplayAlerts = False

    Application.Calculate
    Run callMethod

    MsgBox "YELLO"

    'Capture Waterfall screenshot




End Function

Function getWarpFilePath()
    WarpFilePath = GetXMLData(xmlPath, "//Config/codeRedAlg/warpFilePath")
    WarpTabName = GetXMLData(xmlPath, "//Config/codeRedAlg/warpFilePath", "tabName")

    getWarpFilePath = WarpFilePath
End Function

'@TODO: create a cache version as not not have to always pull from file
Function getCalendarDict()
    Dim gdsBook As Excel.Workbook
    Dim CalBook As Excel.Workbook
    Dim CalDict As Dictionary
    Dim DateDict As Dictionary
    Dim cel As Range

    'Get Config Data
    CalFilePath = GetXMLData(xmlPath, "//Config/codeRedAlg/calendarFilePath")

    Set gdsBook = Application.ActiveWorkbook
    Set CalBook = Workbooks.Open(CalFilePath)
    Set CalDict = New Dictionary

    CalBook.Sheets("Calendar").Activate

    'Cycle through column B and match with column H,I,J
    lastRow = Range("B2").End(xlDown).Row
    For Each cel In Range("B2:B" & lastRow).Cells
        Set DateDict = New Dictionary
        wkDate = cel.Value
        fiscalVer = cel.Offset(0, 6).Value
        fiscalMW = cel.Offset(0, 7).Value
        hyperionVer = cel.Offset(0, 8).Value
        DateDict.Add "fiscalVer", fiscalVer
        DateDict.Add "fiscalMW", fiscalMW
        DateDict.Add "hyperionVer", hyperionVer
        CalDict.Add wkDate, DateDict
    Next

    CalBook.Close

    gdsBook.Sheets("DS_Outlook_Report").Activate

    Set getCalendarDict = CalDict

End Function

'@TODO: at some point object and mapping names should be aligned, so for each key in an object - find correspnding mapping
Function getTemplateMapping()
    Dim TemplateMapping As Dictionary
    Set TemplateMapping = New Dictionary

    'Pull Data From Config
    TemplateMapping.Add "tab_name", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/tab_name")
    TemplateMapping.Add "pn", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/pn")
    TemplateMapping.Add "lob", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/lob")
    TemplateMapping.Add "commodity", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/commodity")
    TemplateMapping.Add "platform", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/platform")
    TemplateMapping.Add "region", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/region")
    TemplateMapping.Add "alert_level", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/alert_level")
    TemplateMapping.Add "recover_from_stbl", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/recover_from_stbl")
    TemplateMapping.Add "recover_from_dsi", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/recover_from_dsi")
    TemplateMapping.Add "current_supply", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/current_supply")
    TemplateMapping.Add "gds", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/gds")
    TemplateMapping.Add "waterfall_warp", GetXMLData(xmlPath, "//Config/codeRedAlg/templateMapping/waterfall_warp")

    Set getTemplateMapping = TemplateMapping
End Function

Function AlertStrToDict()
    Dim parseStr As String
    Dim delimiterRecord As String, delimiterValue As String
    Dim RecordDict As Dictionary, itemDict As Dictionary
    Set RecordDict = New Dictionary

    'Pull Data From Config
    delimiterRecord = GetXMLData(xmlPath, "//Config/delimiters/record")
    delimiterValue = GetXMLData(xmlPath, "//Config/delimiters/value")
    parseStr = GetXMLData(xmlPath, "//Config/codeRedAlg/cache/data")

    Records = Split(parseStr, delimiterRecord)

    For i = LBound(Records) To UBound(Records)
      Record = Records(i)

      RecordVals = Split(Record, delimiterValue)

      'Debug.Print ("Part: " & RecordVals(1) & " | Status: " & RecordVals(3) & " [" & RecordVals(0) & "]")

      partName = RecordVals(1)
      Set itemDict = New Dictionary
      itemDict.Add "Status", RecordVals(3)
      itemDict.Add "Commodity", RecordVals(2)
      'TODO: CHANGE from static value to var for Platform
      itemDict.Add "Platform", "TBD"
      itemDict.Add "Position", RecordVals(0)
      itemDict.Add "Header", RecordVals(4)

      RecordDict.Add partName, itemDict

    Next i

    Set AlertStrToDict = RecordDict

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
