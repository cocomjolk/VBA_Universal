

Sub Cell_Offset()
    'Cell_Offset_Right
    ActiveCell.Offset(, 1).Select
    'Cell_Offset_Left
    ActiveCell.Offset(, -1).Select
    'Cell_Offset_Up
    ActiveCell.Offset(1,).Select
    'Cell_Offset_Down
    ActiveCell.Offset(-1,).Select
End Sub
'--------------------------------------------------------------'

Sub Control_Move()
    'Move to end of column Up
    Selection.End(xlUp).Select
    'Move to end of column Down
    Selection.End(xlDown).Select
    'Move to end of row Right
    Selection.End(xlToRight).Select
    'Move to end of row Left
    Selection.End(xlToLeft).Select
End Sub
'--------------------------------------------------------------'

Sub Change_Column_Width()
    Columns("A:A").ColumnWidth = 9
End Sub
'--------------------------------------------------------------'
Sub Select_Range(rng)

    ActiveCell.Range(rng).Select

End Sub
'--------------------------------------------------------------'

Sub CopyRange(rng)
    Range(rng).Copy
End Sub
'--------------------------------------------------------------'

Sub CopyPaste()

    'Copying hard coded range and pasting into same sheet'
    Range("B2:H8").Select
    Selection.Copy
    Range("J2").Select
    ActiveSheet.Paste

End Sub
'--------------------------------------------------------------'

sub paste_just_values()

  Range("A1").PasteSpecial xlPasteValues

end sub
'--------------------------------------------------------------'
Sub Paste(rng)

    Range(rng).Select
    ActiveSheet.Paste

End Sub
'--------------------------------------------------------------'

Sub SaveAs(path)

    ActiveWorkbook.SaveAs fileName:=path

End Sub
'--------------------------------------------------------------'

Sub copy_paste_to_another_workbook()

    Dim wb1 As Workbook
    Dim wb2 As Workbook

    Set wb1 = Workbooks("WB1.xlsx")
    Set wb2 = Workbooks("WB2.xlsm")

    wb1.Activate
    Range("B2:C5").Select
    Selection.Copy
    wb2.Activate
    Range("B3").Select
    ActiveSheet.Paste

End Sub
'--------------------------------------------------------------'

Sub copy_paste_to_another_sheet()

    Range("B2:H8").Select
    Selection.Copy
    Range("J2").Select
    activeSheet.Paste

    Columns("J:J").ColumnWidth = 19

End Sub
'--------------------------------------------------------------'

sub time_delay()
  Application.Wait Now + #12:00:02 AM#
end sub
'--------------------------------------------------------------'

Sub valueFromCell()

Dim n As Integer
n = Range("A1").Value

MsgBox n

End Sub
'--------------------------------------------------------------'

Sub duplicateSheet()
  ' duplicates sheet1 and places it after sheet1
  Sheets(1).Select
  Sheets(1).Select.Copy after:=Sheets(1)

End Sub
'--------------------------------------------------------------'

Sub deleteSheet()

  Application.DisplayAlerts = False
  activeSheet.Delete
  Application.DisplayAlerts = True

End Sub
'--------------------------------------------------------------'

Sub Get_Cell_Color_From_Cell()
    Dim color As Integer
    color = Range("C4").Interior.color
    MsgBox "Color code is:  " & (color)
End Sub
'--------------------------------------------------------------'

Sub get_workbook_name()

  Dim wkBookName As String
  wkBookName = ActiveWorkbook.Name
  Debug.Print wkBookName
  'MsgBox wkBookName
End Sub
'--------------------------------------------------------------'

Sub get_worksheet_name()

  Dim wkSheetName As String
  wkSheetName = activeSheet.Name
  MsgBox wkSheetName
End Sub
'--------------------------------------------------------------'

Sub get_File_Path()

  Dim FilePath As String

  FilePath = ActiveWorkbook.FullName
  Debug.Print ActiveWorkbook.FullName
  'MsgBox wkBookPath
End Sub
'--------------------------------------------------------------'

Sub get_Folder_Path()

  Dim FolderPath As String

  FolderPath = ActiveWorkbook.Path
  Debug.Print ActiveWorkbook.Path
  'MsgBox FolderPath
End Sub
'--------------------------------------------------------------'

Sub reference_workbook()
  'puts "some value into cell A3 of sheet 1 in workbook.xlsm"'

  Workbooks("workbook.xlsm").Sheets(1).Range("A3") = "some value"
  Workbooks("workbook1.xlsm").Sheets(1).Range("A3").Copy _
  Workbooks("workbook2.xlsm").Sheets(1).Range("A3")

End Sub
'--------------------------------------------------------------'

Sub reference_worksheet()
  worksheets("wsName").Range("A3").Value = "Some value"
  'reference with codename. Good if worksheet names are changed in the future'
  Template1Row.Select
End Sub
'--------------------------------------------------------------'

Sub open_workbook()

  Application.Workbooks.Open "C:\some\file\path.xlsm"

End Sub
'--------------------------------------------------------------'

Sub close_Workbook()

    ActiveWorkbook.Close False

End Sub
'--------------------------------------------------------------'

Sub change_font()

    Dim myRange As Range
    Set myRange = Range("A10", "A" & Cells(Rows.Count, 1).End(xlUp).Row)
    With myRange.Font
        .Name = "Arial"
        .Size = 14
        .Bold = True
    End With

End Sub
'--------------------------------------------------------------'

Sub reset_font()

    Dim myRange As Range
    Set myRange = Range("A10", "A" & Cells(Rows.Count, 1).End(xlUp).Row)
    With myRange.Font
        .Name = "Calibri"
        .Size = 11
        .Bold = False
    End With

End Sub
'--------------------------------------------------------------'

Sub create_custom_collection()

    Dim myCollection As Collection

    Set myCollection = New Collection

    myCollection.Add "some val 1", Value1

End Sub
'--------------------------------------------------------------'

Sub get_ppt_Path()

  Set pptApp = GetObject(, "PowerPoint.Application")
  FilePath = pptApp.ActivePresentation.FullName

End Sub
'--------------------------------------------------------------'

Sub open_powerpoint()

    Dim FilePath   As String

    Set pptApp = CreateObject(class:="PowerPoint.Application")
    pptApp.Presentations.Open (FilePath) ' this works without above code

End Sub
'--------------------------------------------------------------'

Sub open_powerpoint2()

    Dim FilePath  As String

    Set pptApp = CreateObject(class:="PowerPoint.Application").Presentations.Open(FilePath)
    'Set pptPresentation = pptApp.Presentations.Open(FilePath)
End Sub
'--------------------------------------------------------------'



Sub array_one_dimension()

  Dim monthArr(1 to 12) as string
  monthArr(1) = "Jan"
  monthArr(2) = "Feb"

End Sub
'--------------------------------------------------------------'

Sub forLoopList()

  Dim counter As Integer

  For counter = 0 To 10

    'set cell value to counter value
    Selection.Value = counter

    'moves cell down but not left or right
    Selection.Offset(1, 0).Select
  Next counter

End Sub
'--------------------------------------------------------------'

Sub loop_cells_in_range()

    Dim cell As Range

    For Each cell In ActiveSheet.UsedRange
      'do something
    Next cell

End Sub
'--------------------------------------------------------------'

Sub loop_Sheets()

  'This function loops through all worksheets looking for
  'value 43 in cell A1
  'ActiveWorkbook.sheets is like a directory of sheets

  Dim Sh As Worksheet
  Application.DisplayAlerts = False

  'for loop
  For Each Sh In ActiveWorkbook.Sheets
      Dim n As Integer
      n = Sh.Range("A1").Value
      If n <> 43 Then
      MsgBox "Cell A1 does not contain 43" & " Contains: " & n & " in " & Sh.Name
      Else
      MsgBox "Cell A1 " & Sh.Name & " does contain 43 "
      End If
  Next Sh

  Application.DisplayAlerts = True

End Sub
'--------------------------------------------------------------'

Sub tableValues()

  'Dim mySheet As Worksheet

  'Application.DisplayAlerts = False

  'for loop
  'For Each mySheet In ActiveWorkbook.Sheets

    Dim counter As Integer

    Sheets("Sheet2").Copy after:=Sheets("Sheet2") 'copies active sheet and creates new sheet after sheet(3)
    'activeSheet.Name = "New Sheet"
    activeSheet.Range("C3").Value = Sheets("Thermal dashboard").Range("C3")
  'Next counter

  'Application.DisplayAlerts = True

End Sub
'--------------------------------------------------------------'

Sub ExcelRange_to_PPT_Table()

  Dim ppApp As PowerPoint.Application
  Dim ppPres As PowerPoint.Presentation
  Dim ppTbl As PowerPoint.Shape


  On Error Resume Next
  Set ppApp = GetObject(, "PowerPoint.Application")
  On Error GoTo 0

  If ppApp Is Nothing Then
      Set ppApp = New PowerPoint.Application
      Set ppPres = ppApp.Presentations.Item(1)
  Else
      Set ppPres = ppApp.Presentations.Item(1)
  End If

  ppApp.ActivePresentation.Slides(1).Select
  ppPres.Windows(1).Activate

  ' find on Slide Number 1 which object ID is of Table type (you can change to whatever slide number you have your table)
  With ppApp.ActivePresentation.Slides(1).Shapes
      For i = 1 To .Count
          If .Item(i).HasTable Then
              ShapeNum = i
          End If
      Next
  End With

  ' assign Slide Table object
  Set ppTbl = ppApp.ActivePresentation.Slides(1).Shapes(ShapeNum)

  ' copy range from Excel sheet
  iLastRowReport = Range("B" & Rows.Count).End(xlUp).Row
  Range("A3:J" & iLastRowReport).Copy

  ' select the Table cell you want to copy to >> modify according to the cell you want to use as the first Cell
  ppTbl.Table.Cell(3, 1).Shape.Select

  ' paste into existing PowerPoint table - use this line if you want to use the PowerPoint table format
  ppApp.CommandBars.ExecuteMso ("PasteExcelTableDestinationTableStyle")

  ' paste into existing PowerPoint table - use this line if you want to use the Excel Range format
  ' ppApp.CommandBars.ExecuteMso ("PasteExcelChartSourceFormatting")

End Sub
'--------------------------------------------------------------'

Sub rangeTest()
    'Dim C4 As Integer

'    Worksheets("Template (6)").Select
'    C4 = Range("C4").Value
'    Worksheets("Template").Select
'    Range("C4") = C4

    'The below code does same thing as AboveAverage code
    Sheets("Template").Range("C4").Value = Sheets("Template (6)").Range("C4")


    'Range("C4:H8").Cells(1, 1) = 1
    'MsgBox Sheets("Template").Range("C4")

End Sub
'--------------------------------------------------------------'

Sub create_newbook()

  Dim newBook as Workbook
  Dim newSheet as Worksheets
  Dim newRange as range

  newSheet.Range("A1").Value = "some value"
  'thisWorkbook referred to the workbook on top'
  Set newRange = ThisWorkbook.Worksheets("Dim").Range("B5:D5")
  newRange.Font.Bold = True
  newRange.Interior.Color = VBA.RGB(255, 255, 204)

End Sub
'--------------------------------------------------------------'

'To find the end of a column of rows can use range("A10","A"& cells(rows.Count,1).end(xlup).row)...'
'to go from cell "A10" to last row of data in column A'

Set rng = Range("A:A").Find(what:=Range("B3").Value,LookIn:=xlValues, lookat:=xlWhole)
range("C3").Value=compid.Offset(,4)
'--------------------------------------------------------------'

Sub RefreshAll()

    ActiveWorkbook.RefreshAll

End Sub
'--------------------------------------------------------------'

Sub FillSeries(rng)
    Range(rng).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Date:=xlDay, _
        Step:=1, Trend:=False
End Sub
'--------------------------------------------------------------'

Sub InsertLeft(rng)

    Range(rng).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

End Sub
'--------------------------------------------------------------'

Sub WiggleCell()

    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select

End Sub
'--------------------------------------------------------------'

Sub ClearRange(rng)

    Range(rng).Clear

End Sub
'--------------------------------------------------------------'

Sub DeleteRange(rng)

    Range(rng).Delete

End Sub
'--------------------------------------------------------------'

Sub UnMerge(val)

    Range(val).UnMerge

End Sub
'--------------------------------------------------------------'

Sub Todays_date()

    ActiveCell.value = Date
    ActiveCell.Offset(1, 0).Range("A1").Select

End Sub
'--------------------------------------------------------------'

Sub Add_To_Date()

    'DateAdd ( "interval", number, date )

    ActiveCell.value = DateAdd("d", 1, Date)
    ActiveCell.Offset(1, 0).Range("A1").Select

End Sub
'--------------------------------------------------------------'

Sub Highlight_Range(rng)

    Range(rng).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub
'--------------------------------------------------------------'

Sub DragDown(rng1, rng2)

    ' highlights the cell row currently on, then drags down.
    ' rng1 selects starting range (relative) ex. "A1:Q1"
    ' rng2 drags down to row number desired ex. "A1:Q250"
    ' will drag down 250 rows from starting cell (relative).
    ActiveCell.Range(rng1).Select
    Selection.AutoFill Destination:=ActiveCell.Range(rng2), Type:= _
        xlFillDefault

End Sub
'--------------------------------------------------------------'

Sub SaveAs_Xlsb(path As String, Optional format As Variant = 51)

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs fileName:=path, FileFormat:=format
    Application.DisplayAlerts = True

End Sub
'--------------------------------------------------------------'

Sub FilterOr(val1, val2)

   ActiveSheet.Range("$A$1").AutoFilter field:=ActiveCell.Column, Criteria1:=val1 _
       , Operator:=xlOr, Criteria2:=val2
End Sub
'--------------------------------------------------------------'

Sub AutomaticCalculations()

    Application.Calculation = xlAutomatic

End Sub
'--------------------------------------------------------------'

Sub ManualCalculations()

    Application.Calculation = xlManual

End Sub
'--------------------------------------------------------------'

Sub ClearFilters()

    If ActiveSheet.AutoFilterMode Then Cells.AutoFilter

End Sub

'--------------------------------------------------------------'

Sub FilterColumn(criteria)

    ActiveSheet.Range(ActiveCell.Address).AutoFilter field:=13, Criteria1:= _
        criteria, Operator:=xlAnd

End Sub
'--------------------------------------------------------------'

Sub FilterPivotField(table, field, value)

    ' FilterPivotField Macro
    ActiveSheet.PivotTables(table).PivotFields(field). _
        EnableMultiplePageItems = False
    ActiveSheet.PivotTables("PivotTable2").PivotFields("OC_Type").CurrentPage = _
        value
End Sub
'--------------------------------------------------------------'

Sub SortDescending(rng)
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range(rng)
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
'--------------------------------------------------------------'

Sub FindFileName()

    Dim strFileName As String
    Dim FileList() As String
    Dim intFoundFiles As Integer
    Dim activeGdsBook As Workbook

    WriteToLog "==========================================="
    WriteToLog "STARTING FULL PROCESS"

    fileFolder = GetXMLData(xmlPath, "//Config/codeRedAlg/inputPath")

    Dim strFolder As String: strFolder = fileFolder
    Dim strFileSpec As String: strFileSpec = strFolder & "*.*"
    strFileName = Dir(strFileSpec)
    Do While Len(strFileName) > 0

        ReDim Preserve FileList(intFoundFiles)

        'Check to make sure to skip any files with bad file name characters
        If ((InStr(1, strFileName, "?") <> 0) <> True) Then
            'For all GOOD files, run process for it

            WriteToLog "FILE: " & fileFolder & strFileName
            'Run methods

        End If
        FileList(intFoundFiles) = strFileName
        intFoundFiles = intFoundFiles + 1
        strFileName = Dir
    Loop

    WriteToLog "CRA Data"

    WriteToLog "=============================================="

    Debug.Print "FILE: " & strFileName

    WriteToLog "END FULL PROCESS"

End Sub

'--------------------------------------------------------------'

Function FindFileName()

    Dim strFileName As String
    Dim FileList() As String
    Dim intFoundFiles As Integer
    Dim activeGdsBook As Workbook

    fileFolder = "Z:\SN1213 PPT Creation\GDS\"

    Dim strFolder As String: strFolder = fileFolder
    Dim strFileSpec As String: strFileSpec = strFolder & "*.*"
    strFileName = Dir(strFileSpec)
    Do While Len(strFileName) > 0

        ReDim Preserve FileList(intFoundFiles)

        'Check to make sure to skip any files with bad file name characters
        If ((InStr(1, strFileName, "?") <> 0) <> True) Then
            'For all GOOD files, run process for it
            'Run methods
            Debug.Print "FILE: " & strFileName

        End If
        FileList(intFoundFiles) = strFileName
        intFoundFiles = intFoundFiles + 1
        strFileName = Dir
    Loop

End Function
'--------------------------------------------------------------'

Sub Dictionary_Late_Binding()
    'In technical terms Early binding means we decide exactly
    'what we are using up front. With Late binding this decision
    'is made when the application is running. In simple terms the difference is.
    '1. Early binding requires a reference. Late binding doesn’t.
    '2. Early binding allows access to *Intellisense. Late binding doesn’t.
    '3. Early binding may require you to manually add the Reference to the “Microsoft Scripting Runtime” for some users.
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

End Sub
'--------------------------------------------------------------'

Sub Declare_A_Dictionary()

    Dim dict As New Scripting.Dictionary
    'Or use, both do the same thing'
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary

End Sub
'--------------------------------------------------------------'

Sub Sub_name()

    Dim dict As New Scripting.Dictionary

    ' Add to fruit to Dictionary
    dict.Add key:="Apple", Item:=51
    dict.Add key:="Peach", Item:=34
    dict.Add key:="Plum", Item:=43

    Dim sFruit As String
    ' Ask user to enter fruit
    sFruit = InputBox("Please enter the name of a fruit")

    If dict.Exists(sFruit) Then
      MsgBox sFruit & " exists and has value " & dict(sFruit)
    Else
      MsgBox sFruit & " does not exist."
    End If

    Set dict = Nothing

End Sub
'--------------------------------------------------------------'

Sub Sub_name()
    Dim dict As New Scripting.Dictionary

    ' Add some items
    dict.Add "Orange", 55
    dict.Add "Peach", 55
    dict.Add "Plum", 55
    Debug.Print "The number of items is " & dict.Count

    ' Remove one item
    dict.Remove "Orange"
    Debug.Print "The number of items is " & dict.Count

    ' Remove all items
    dict.RemoveAll
    Debug.Print "The number of items is " & dict.Count
End Sub
'--------------------------------------------------------------'
Sub Sub_name()
    'code here'
End Sub

Sub WriteToArr_ReadFromArr()

    i = 1
    Set ExceptionList = New Dictionary

    Do While i < 5
        'Add to array
           ExceptionList.Add "PNABC" & i, "Some value __ " & i
        i = i + 1
    Loop

    For Each exceptionPart In ExceptionList.Keys()
        Debug.Print exceptionPart, ExceptionList.Item(exceptionPart)
    Next

End Sub
'--------------------------------------------------------------'
'--------------------------------------------------------------'
'--------------------------------------------------------------'
'--------------------------------------------------------------'

