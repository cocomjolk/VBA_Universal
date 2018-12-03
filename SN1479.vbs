Sub test()

    Dim wsGDS As String
    Dim wbGDS As String
    Dim fileName As String

    'get file from shared drive
    Call FindFileName




    Application.Workbooks.Open ("Z:\SN1213 PPT Creation\GDS\" & GDSFileName)
    wbGDS = ActiveWorkbook.Name
    wsGDS = ActiveSheet.Name

    Call Activate_GDS_Sheet(wbGDS, wsGDS)

    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$AN$1").AutoFilter field:=3, Criteria1:="WW"
    ActiveSheet.Range("$A$1:$AN$1").AutoFilter field:=7, Criteria1:= _
    "Projected DSI"

    ActiveWorkbook.Sheets.Add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "Exceptions"
    Call Activate_GDS_Sheet(wbGDS, wsGDS)

End Sub


Sub Create_Weekly_GDS()



  'create tab in GDS file expectations'

  'Set sort filter
  temp = Activate_GDS_Sheet()
  Rows("1:1").Select
  Selection.AutoFilter
  ActiveSheet.Range("$A$1:$AM$281").AutoFilter Field:=3, Criteria1:="WW"
  ActiveSheet.Range("$A$1:$AM$281").AutoFilter Field:=6, Criteria1:= _
  "Projected DSI"

  'get array of PN
  Range("A1").Select
  ActiveCell.Offset(10, 0).Range("A1").Select

  'Copy dates from weekly GDS to paste
  Range("H1").Select
  Range(Selection, Selection.End(xlToRight)).Select
  Selection.Copy

  'paste dates to Dashboard
  temp = Activate_Dashboard_Sheet()
  Range("F2").Select
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

'loop do while PN2 <> "" '
  Do While PN2 <> ""

        'get PN1Match row and PN2Match row range
        temp = Activate_Dashboard_Sheet()
        Range("D2").Select

        'if PN1 deos not equal active cell go down else stop
        Do While PN1 <> PN1Match
          PN1Match = ActiveCell.value
          PN1MatchRow = ActiveCell.Row
          ActiveCell.Offset(1, 0).Range("a1").Select
        Loop
        Do While PN2 <> PN2Match
          PN2Match = ActiveCell.value
          PN2MatchRow = ActiveCell.Row
          ActiveCell.Offset(1, 0).Range("a1").Select
        Loop

        PNCounter = PN2MatchRow - PN1MatchRow
        'Debug.Print PNCounter

        For i = 1 To PNCounter
          'Cells(i, 1).value = 100
          temp = Activate_GDS_Sheet()
          Range("H" & PN1Row).Select
          Range(Selection, Selection.End(xlToRight)).Select
          Selection.Copy
          temp = Activate_Dashboard_Sheet()
          Range("F" & PN1MatchRow).Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            PN1MatchRow = PN1MatchRow + 1
        Next i

        temp = Activate_GDS_Sheet()
        Range("A" & PN1Row ).Select
        ActiveCell.Offset(10, 0).Range("A1").Select
        PN1 = ActiveCell.Value
        PN1Row = ActiveCell.Row
        ActiveCell.Offset(10, 0).Range("A1").Select
        PN2 = ActiveCell.Value
        PN2Row = ActiveCell.Row

  Loop

End Sub

Sub Activate_GDS_Sheet()
  Windows("Tool_Output.xlsx").Activate
  Sheets("Weekly GDS").Select
End Sub

Sub Activate_Dashboard_Sheet()
  Windows("BETA_NB_PCB_DSI Dashboard Template.xlsx").Activate
  Sheets("Original").Select
End Sub

Sub test()
  Dim PN1 As String
  Dim PN2 As String
  Dim PN1Match As String
  Dim PN2Match As String
  Dim PN1Row As Integer
  Dim PN2Row As Integer
  Dim PN1MatchRow As Integer
  Dim PN2MatchRow As Integer
  Dim PNCounter As Integer


  PN1 = "CW03M"
  PN2 = "369FR"


  temp = Activate_Dashboard_Sheet()
  Range("D2").Select

  'if PN1 deos not equal active cell go down else stop
  Do While PN1 <> PN1Match
    PN1Match = ActiveCell.value
    PN1MatchRow = ActiveCell.Row
    ActiveCell.Offset(1, 0).Range("a1").Select
  Loop

  Do While PN2 <> PN2Match
    PN2Match = ActiveCell.value
    PN2MatchRow = ActiveCell.Row
    ActiveCell.Offset(1, 0).Range("a1").Select
  Loop

  PNCounter = PN2MatchRow - PN1MatchRow
  'Debug.Print PNCounter

  For i = 1 To PNCounter
    'Cells(i, 1).value = 100
    temp = Activate_GDS_Sheet()
    Range("H" & PN1Row).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    temp = Activate_Dashboard_Sheet()
    Range("F" & PN1MatchRow).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
      PN1MatchRow = PN1MatchRow + 1
  Next i


End Sub

Public log As TextStream
Public logPath As String
Public xmlPath As String
Sub setXmlPath(Path)
    xmlPath = Path
End Sub
Sub setLogPath(Path)
    logPath = Path
    Debug.Print logPath
End Sub

Sub test()
   temp = FindFileName()
End Sub

Sub Create_Weekly_GDS()

  Dim PN1 As String
  Dim PN2 As String
  Dim PN1Match As String
  Dim PN2Match As String
  Dim PN1Row As Integer
  Dim PN2Row As Integer
  Dim PN1MatchRow As Integer
  Dim PN2MatchRow As Integer


  'Set sort filter
  temp = Activate_GDS_Sheet()
  Rows("1:1").Select
  Selection.AutoFilter
  ActiveSheet.Range("$A$1:$AM$281").AutoFilter field:=3, Criteria1:="WW"
  ActiveSheet.Range("$A$1:$AM$281").AutoFilter field:=6, Criteria1:= _
  "Projected DSI"

  'get PN1 and PN2 first time'
  Range("A1").Select
  ActiveCell.Offset(10, 0).Range("A1").Select
  PN1 = ActiveCell.value
  PN1Row = ActiveCell.Row
  ActiveCell.Offset(10, 0).Range("A1").Select
  PN2 = ActiveCell.value
  PN2Row = ActiveCell.Row

  'Copy dates from weekly GDS to paste
  Range("H1").Select
  Range(Selection, Selection.End(xlToRight)).Select
  Selection.Copy

  'paste dates to Dashboard
  temp = Activate_Dashboard_Sheet()
  Range("F2").Select
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False

'loop do while PN2 <> "" '
  Do While PN2 <> ""

        'get PN1Match row and PN2Match row range
        temp = Activate_Dashboard_Sheet()
        Range("D2").Select

        'if PN1 deos not equal active cell go down else stop
        Do While PN1 <> PN1Match
          PN1Match = ActiveCell.value
          PN1MatchRow = ActiveCell.Row
          ActiveCell.Offset(1, 0).Range("a1").Select
        Loop
        Do While PN2 <> PN2Match
          PN2Match = ActiveCell.value
          PN2MatchRow = ActiveCell.Row
          ActiveCell.Offset(1, 0).Range("a1").Select
        Loop

        PNCounter = PN2MatchRow - PN1MatchRow
        'Debug.Print PNCounter

        For i = 1 To PNCounter
          'Cells(i, 1).value = 100
          temp = Activate_GDS_Sheet()
          Range("H" & PN1Row).Select
          Range(Selection, Selection.End(xlToRight)).Select
          Selection.Copy
          temp = Activate_Dashboard_Sheet()
          Range("F" & PN1MatchRow).Select
          Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
              :=False, Transpose:=False
            PN1MatchRow = PN1MatchRow + 1
        Next i

        temp = Activate_GDS_Sheet()
        Range("A" & PN1Row).Select
        ActiveCell.Offset(10, 0).Range("A1").Select
        PN1 = ActiveCell.value
        PN1Row = ActiveCell.Row
        ActiveCell.Offset(10, 0).Range("A1").Select
        PN2 = ActiveCell.value
        PN2Row = ActiveCell.Row

  Loop

End Sub

Function Activate_GDS_Sheet()
  Windows("Tool_Output.xlsx").Activate
  Sheets("Weekly GDS").Select
End Function

Function Activate_Dashboard_Sheet()
  Windows("BETA_NB_PCB_DSI Dashboard Template.xlsx").Activate
  Sheets("Original").Select
End Function

Function FindFileName()

    Dim strFileName As String
    Dim FileList() As String
    Dim intFoundFiles As Integer
    Dim activeGdsBook As Workbook

    fileFolder = "Z:\SN1213 PPT Creation\Input Files\Power Input Files\"

    Dim strFolder As String: strFolder = fileFolder
    Dim strFileSpec As String: strFileSpec = strFolder & "*.*"
    strFileName = Dir(strFileSpec)
    Do While Len(strFileName) > 0

        ReDim Preserve FileList(intFoundFiles)

        'Check to make sure to skip any files with bad file name characters
        If ((InStr(1, strFileName, "?") <> 0) <> True) Then
            'For all GOOD files, run process for it

            Debug.Print "FILE NAME: "; strFileName
            'Run methods

        End If
        FileList(intFoundFiles) = strFileName
        intFoundFiles = intFoundFiles + 1
        strFileName = Dir
    Loop

End Function
