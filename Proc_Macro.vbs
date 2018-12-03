'FRED_WU1 2018/8/10
Option Explicit
Sub Main_GDS_Daily_Integrated()
Daily_Weekly = "Daily"
Application.ScreenUpdating = False
Call collect_Checkbox_data
Call Get_Foxjua_Inv

Dim Class_PrepareData As Object
Set Class_PrepareData = New Cls_PrepareData
Class_PrepareData.Arr_Measure = Array("Released Forecast", "Backlog", "Future Backlog", "Net Supply", "Inventory On-Hand")
Class_PrepareData.Arr_Measure_Min = Array("Released Forecast", "Backlog", "Future Backlog", "Net Supply", "Inventory On-Hand", "SLC On-Hand")

Class_PrepareData.Main_Prepare_Data
Arr_PNGroup = Class_PrepareData.Arr_PNGroup
Arr_SiteGroup = Class_PrepareData.Arr_SiteGroup
Arr_pitdata = Class_PrepareData.Arr_pitdata
ExistWW = Class_PrepareData.ExistWW
Arr_StartColumnDateWK = Class_PrepareData.Arr_StartColumnDateWK
Set Class_PrepareData = Nothing
If Supplier_Senstive = False Then
    Call Get_Dic_Index_Pitdata
Else: Call Get_Dic_Index_Pitdata_SupplierSenstive
End If

Call Get_Output_FistRow
Call Get_Arr_Formula_DSI
If Supplier_Senstive = False Then
    Call Get_CoreData
Else: Call Get_CoreData_SupplierSenstive
End If
If UBound(Arr_output, 2) = 1 Then MsgBox ("No data was presented ,please check whether Pitdata/Setting is correct!"): End
Workbooks.Add
Do While Sheets.Count > 1
    Sheets(Sheets.Count).Delete
Loop
ActiveSheet.Range(Cells(1, 1), Cells(UBound(Arr_output, 2), 1)).NumberFormat = "@"
ActiveSheet.Name = "Daily GDS Integrated"
Cells(1, 1).Resize(UBound(Arr_output, 2), UBound(Arr_output)) = Application.Transpose(Arr_output)
Dim Class_Format As Object
Set Class_Format = New Cls_Format
Call Class_Format.Formating_GDS
Call Class_Format.Formating_GDS_DailyIntegrated
Application.ScreenUpdating = True
Call ClearRAM
End Sub
Private Sub Get_Output_FistRow()
Dim Arr_Title, A%, B%
If Supplier_Senstive = False Then
    Arr_Title = Array("Part-Group", "Description", "Site-Group", "CFG", "Buyer Name", "Measure", "BLG/INV", Arr_StartColumnDateWK(2) - 1)
Else: Arr_Title = Array("Part-Group", "Description", "Site-Group", "Supplier", "CFG", "Buyer Name", "Measure", "BLG/INV", Arr_StartColumnDateWK(2) - 1)
End If
For A = 9 To UBound(Arr_pitdata, 1)
    If CDate(Arr_pitdata(A, 1)) >= CDate(Arr_StartColumnDateWK(2)) Then
        ReDim Preserve Arr_Title(0 To UBound(Arr_Title) + 1)
        Arr_Title(UBound(Arr_Title)) = CDate(Arr_pitdata(A, 1))
    End If
Next A

ReDim Arr_output(1 To UBound(Arr_Title) + 1, 1 To 1)
For B = 1 To UBound(Arr_Title) + 1
    Arr_output(B, 1) = Arr_Title(B - 1)
Next B
End Sub
Private Sub Get_Arr_Formula_DSI()
ReDim Arr_Formula_DSI(1 To UBound(Arr_output, 1))
Dim Arr_ColumnStart_endDSI
Dim A%
For A = 7 To UBound(Arr_output) - 4
    If IsDate(Arr_output(A, 1)) = True Then
        Arr_ColumnStart_endDSI = Get_Arr_ColumnStart_endDSI(Arr_output(A, 1))
        If IsArray(Arr_ColumnStart_endDSI) Then
        Arr_Formula_DSI(A) = "=IFERROR(R[-1]C/(SUM(R[-4]C" & Arr_ColumnStart_endDSI(1) & ":R[-4]C" & Arr_ColumnStart_endDSI(2) & ")/20),0)"
        End If
    End If
Next A
End Sub
Private Function Get_Arr_ColumnStart_endDSI(InputDate)
ReDim Arr_ColumnStart_endDSI(1 To 2)
Dim Start_wk As Date
Start_wk = Get_StartWKfromDate(InputDate)
Dim A%
For A = 7 To UBound(Arr_output)
    If Arr_output(A, 1) = Start_wk + 7 Then
        Arr_ColumnStart_endDSI(1) = A
    ElseIf Arr_output(A, 1) = Start_wk + 28 Then
        Arr_ColumnStart_endDSI(2) = A
        Get_Arr_ColumnStart_endDSI = Arr_ColumnStart_endDSI
        Exit Function
    End If
Next A
End Function

'FRED_WU1 2018/8/10
Option Explicit
Dim Need_Subtotal As Boolean, NeedWW As Boolean
Sub Main_GDS_Daily()
Application.ScreenUpdating = False
Call collect_Checkbox_data

Dim Class_PrepareData As Object
Set Class_PrepareData = New Cls_PrepareData
Class_PrepareData.Arr_Measure = Array("Released Forecast", "Future Backlog", "Backlog", "Average Forecast", "Commit Air", "Commit Surface", "ASN Air", "ASN Surface", "Receipts", "Net Supply", "SLC On-Hand", "Factory On-Hand", "Inventory On-Hand", "Projected Inventory", "Projected DSI")
Class_PrepareData.Arr_Measure_Min = Array("Released Forecast", "Future Backlog", "Average Forecast", "Net Supply", "Projected Inventory", "Projected DSI", "SLC On-Hand")

Class_PrepareData.Main_Prepare_Data
Arr_PNGroup = Class_PrepareData.Arr_PNGroup
Arr_SiteGroup = Class_PrepareData.Arr_SiteGroup
Arr_pitdata = Class_PrepareData.Arr_pitdata
ExistWW = Class_PrepareData.ExistWW
Arr_StartColumnDateWK = Class_PrepareData.Arr_StartColumnDateWK
Set Class_PrepareData = Nothing

If Supplier_Senstive = False Then
    Call Get_Dic_Index_Pitdata
Else: Call Get_Dic_Index_Pitdata_SupplierSenstive
End If
Call Get_Output_FistRow
If Supplier_Senstive = False Then
    Call Get_CoreData_Daily
Else: Call Get_CoreData_Daily_SupplierSensitive
End If

If UBound(Arr_output, 2) = 1 Then MsgBox ("No data was presented ,please check whether Pitdata/Setting is correct!"): End
Call Fill_in_Formula
Call Fill_in_Formula_WW

Workbooks.Add
Do While Sheets.Count > 1
    Sheets(Sheets.Count).Delete
Loop
ActiveSheet.Range(Cells(1, 1), Cells(UBound(Arr_output, 2), 2)).NumberFormat = "@"
ActiveSheet.Name = "Daily GDS"
Cells(1, 1).Resize(UBound(Arr_output, 2), UBound(Arr_output)) = Application.Transpose(Arr_output)
Dim Class_Format As Object
Set Class_Format = New Cls_Format
Call Class_Format.Formating_GDS_Daily(Arr_StartColumnDateWK(2))
Application.ScreenUpdating = True
Call ClearRAM
End Sub
Private Sub Get_CoreData_Daily()
Dim A%, B%
For A = 1 To UBound(Arr_PNGroup)
    For B = 1 To UBound(Arr_SiteGroup)
        Call Get_Daily_GDSData_PNsite(Arr_PNGroup(A, 1), Arr_SiteGroup(B, 1), Arr_PNGroup(A, 2), Arr_SiteGroup(B, 2))
        If Need_Subtotal = True Then
            Call Get_Daily_GDSData_Subtotal(Arr_PNGroup(A, 1), Arr_SiteGroup(B, 1)): Need_Subtotal = False
        End If
    Next B
    If ExistWW = True And NeedWW = True Then
        Call Get_Daily_GDSData_WW(Arr_PNGroup(A, 1), Arr_PNGroup(A, 2)): NeedWW = False
    End If
Next A
End Sub
Private Sub Get_Daily_GDSData_PNsite(PNgroup, Sitegroup, Arr_PNs, Arr_Sites, Optional Supplier As String)
Dim A%, B%, C&, D&, E%, T&, Arr_temp_row()
For A = LBound(Arr_PNs) To UBound(Arr_PNs)
    For B = LBound(Arr_Sites) To UBound(Arr_Sites)
        If Dic_Index_Pitdata.EXISTS(Arr_PNs(A) & Arr_Sites(B) & Supplier) Then
            Need_Subtotal = True: NeedWW = True
            Arr_temp_row = Dic_Index_Pitdata(Arr_PNs(A) & Arr_Sites(B) & Supplier)
            T = UBound(Arr_output, 2)
            For C = LBound(Arr_temp_row) To UBound(Arr_temp_row)
                For D = Arr_temp_row(C) To UBound(Arr_pitdata, 2)
                    If Arr_pitdata(1, D) = Arr_PNs(A) And Arr_pitdata(3, D) = Arr_Sites(B) Then
                        T = T + 1
                        ReDim Preserve Arr_output(1 To UBound(Arr_output), 1 To T)
                        Arr_output(1, T) = Arr_pitdata(1, D)
                        Arr_output(2, T) = PNgroup
                        Arr_output(3, T) = Arr_pitdata(2, D)
                        Arr_output(4, T) = Arr_pitdata(3, D)
                        Arr_output(5, T) = Sitegroup
                        For E = 4 To UBound(Arr_pitdata)
                            Arr_output(E + 2, T) = Arr_pitdata(E, D)
                        Next E
                    Else: Exit For
                    End If
                Next D
            Next C
        End If
    Next B
Next A
End Sub
Private Sub Get_Daily_GDSData_Subtotal(PNgroup, Sitegroup, Optional Supplier As String)
Dim A&, B%, T&
T = UBound(Arr_output, 2)
ReDim Preserve Arr_output(1 To UBound(Arr_output), 1 To T + 5)
For A = T + 1 To T + 5
    Arr_output(2, A) = PNgroup
    If Not Supplier = Empty Then Arr_output(3, A) = Supplier
    Arr_output(5, A) = Sitegroup

    For B = 4 To 7
        Arr_output(B + 2, A) = Arr_output(B + 2, T)
    Next B
Next A
Arr_output(10, T + 1) = "Subtotal Forecast"
Arr_output(10, T + 2) = "Subtotal Future BLG"
Arr_output(10, T + 3) = "Subtotal Supply"
Arr_output(10, T + 4) = "Subtotal Inventory"
Arr_output(10, T + 5) = "Subtotal DSI"
End Sub
Private Sub Get_Daily_GDSData_WW(PNgroup, Arr_PNs, Optional Supplier As String)
Dim A&, B%, T&
T = UBound(Arr_output, 2)
ReDim Preserve Arr_output(1 To UBound(Arr_output), 1 To T + 5)
For A = T + 1 To T + 5
    Arr_output(2, A) = PNgroup
    Arr_output(5, A) = "WW"
    For B = 4 To 7
        Arr_output(B + 2, A) = Arr_output(B + 2, T)
    Next B
Next A
Arr_output(10, T + 1) = "WW Total Forecast"
Arr_output(10, T + 2) = "WW Total Future BLG"
Arr_output(10, T + 3) = "WW Total Supply"
Arr_output(10, T + 4) = "WW Total Inventory"
Arr_output(10, T + 5) = "WW Total DSI"
End Sub
Private Sub Addjust_FSJ_INV()





End Sub
Private Sub Fill_in_Formula()
Dim Criterion_Date As Date
Criterion_Date = Arr_StartColumnDateWK(2)
Dim A%, B%, C%, D%, E%, F%, G%, H%, M%, N%, RFrow%, Frow%, AFrow%, NSrow%, X&, Y%, PR&
For M = 2 To 16
    If Arr_output(10, M) = "Released Forecast" And A = Empty Then A = M
    If Arr_output(10, M) = "Average Forecast" And B = Empty Then B = M
    If Arr_output(10, M) = "Net Supply" And C = Empty Then C = M
    If Arr_output(10, M) = "Future Backlog" And N = Empty Then N = M
    If Arr_output(10, M) = "Projected Inventory" And D = Empty Then D = M
Next
RFrow = D - A: AFrow = D - B + 1: NSrow = D - C: Frow = D - N
For H = 10 To UBound(Arr_output, 1)
    If Arr_output(H, 1) = Criterion_Date Then Y = H + 1: Exit For
Next
For X = 1 To UBound(Arr_output, 2)
    If Arr_output(10, X) = "Projected Inventory" Then
        For E = Y To UBound(Arr_output, 1)
            Arr_output(E, X) = "=SUM(RC[-1],R[-" & NSrow & "]C,-R[-" & RFrow & "]C,-R[-" & Frow & "]C)"
            Arr_output(E, X + 1) = "=IFERROR(R[-1]C/R[-" & AFrow & "]C,"""")"
        Next
    End If
    If Arr_output(10, X) = "Subtotal Forecast" Then
        For F = X - 1 To 2 Step -1
            If Arr_output(1, F) = Empty Then
                PR = X - F - 1: Exit For
            ElseIf F = 2 Then
                PR = X - F: Exit For
            End If
        Next F
        For G = 11 To UBound(Arr_output, 1)
            Arr_output(G, X) = "=IF(SUMIF(R[-" & PR & "]C10:R[-1]C10,""Released Forecast"",R[-" & PR & "]C:R[-1]C)<>0,SUMIF(R[-" & PR & "]C10:R[-1]C10,""Released Forecast"",R[-" & PR & "]C:R[-1]C),"""")"
            Arr_output(G, X + 1) = "=IF(SUMIF(R[-" & PR + 1 & "]C10:R[-2]C10,""Future Backlog"",R[-" & PR + 1 & "]C:R[-2]C)<>0,SUMIF(R[-" & PR + 1 & "]C10:R[-2]C10,""Future Backlog"",R[-" & PR + 1 & "]C:R[-2]C),"""")"
            Arr_output(G, X + 2) = "=IF(SUMIF(R[-" & PR + 2 & "]C10:R[-3]C10,""Net Supply"",R[-" & PR + 2 & "]C:R[-3]C)<>0,SUMIF(R[-" & PR + 2 & "]C10:R[-3]C10,""Net Supply"",R[-" & PR + 2 & "]C:R[-3]C),"""")"

            If G < Y Then
            Arr_output(G, X + 3) = "=IF(SUMIF(R[-" & PR + 3 & "]C10:R[-4]C10,""Projected Inventory"",R[-" & PR + 3 & "]C:R[-4]C)<>0,SUMIF(R[-" & PR + 3 & "]C10:R[-4]C10,""Projected Inventory"",R[-" & PR + 3 & "]C:R[-4]C),"""")"
            Else: Arr_output(G, X + 3) = "=SUM(RC[-1])+SUM(R[-1]C)-SUM(R[-2]C)-SUM(R[-3]C)"
            End If
            Arr_output(G, X + 4) = "=IFERROR(R[-1]C/SUMIF(R[-" & PR + 4 & "]C10:R[-5]C10,""Average Forecast"",R[-" & PR + 4 & "]C:R[-5]C),0)"
        Next G
    End If
LINE1:
Next
End Sub
Private Sub Fill_in_Formula_WW()
Dim Criterion_Date As Date
Criterion_Date = Arr_StartColumnDateWK(2)
Dim A%, B%, C%, D%, E%, F%, StrartTemp%
For A = 10 To UBound(Arr_output, 1)
    If Arr_output(A, 1) = Criterion_Date Then B = A: Exit For
Next
C = UBound(Arr_output, 2) - 3
Do While C > 1
    If Arr_output(5, C) = "WW" And Arr_output(10, C) = "WW Total Forecast" Then
        For D = C - 1 To 1 Step -1
            If Arr_output(5, D) = "WW" And Arr_output(10, D) = "WW Total DSI" Or D = 1 Then
                StrartTemp = D + 1: Exit For
            End If
        Next D
        For E = B To UBound(Arr_output, 1)
            Arr_output(E, C) = "=IF(SUMIF(R" & StrartTemp & "C10:R[-1]C10,""Subtotal Forecast"",R" & StrartTemp & "C:R[-1]C)=0,"""",SUMIF(R" & StrartTemp & "C10:R[-1]C10,""Subtotal Forecast"",R" & StrartTemp & "C:R[-1]C))"
            Arr_output(E, C + 1) = "=IF(SUMIF(R" & StrartTemp & "C10:R[-2]C10,""Subtotal Future BLG"",R" & StrartTemp & "C:R[-2]C)=0,"""",SUMIF(R" & StrartTemp & "C10:R[-2]C10,""Subtotal Future BLG"",R" & StrartTemp & "C:R[-2]C))"
            Arr_output(E, C + 2) = "=IF(SUMIF(R" & StrartTemp & "C10:R[-3]C10,""Subtotal Supply"",R" & StrartTemp & "C:R[-3]C)=0,"""",SUMIF(R" & StrartTemp & "C10:R[-3]C10,""Subtotal Supply"",R" & StrartTemp & "C:R[-3]C))"
            Arr_output(E, C + 4) = "=IFERROR(R[-1]C/SUMIF(R" & StrartTemp & "C10:R[-5]C10,""Average Forecast"",R" & StrartTemp & "C:R[-5]C),""-"")"
'        "=IFERROR(R[-1]C/SUMIF(R2C10:R[-5]C[-4],""Average Forecast"",R2C:R[-5]C),""-"")"
        Next E
        Arr_output(B, C + 3) = "=SUMIF(R" & StrartTemp & "C10:R[-4]C10,""Subtotal Inventory"",R" & StrartTemp & "C:R[-4]C)"
        For F = B + 1 To UBound(Arr_output, 1)
            Arr_output(F, C + 3) = "=SUM(RC[-1])+SUM(R[-1]C)-SUM(R[-2]C)-SUM(R[-3]C)"
        Next F
    End If
C = C - 1
Loop
End Sub
Private Sub Get_Output_FistRow()
ReDim Arr_output(1 To UBound(Arr_pitdata) + 2, 1 To 1)
Dim Arr_Title()
Arr_Title = Array("#Customer Item", "Item Group", "Supplier", "Customer Site", "Regions")
Dim A%
For A = 1 To 5
    Arr_output(A, 1) = Arr_Title(A - 1)
Next A
Dim B%
For B = 6 To 10
    Arr_output(B, 1) = Arr_pitdata(B - 2, 1)
Next B
Dim C%
For C = 11 To UBound(Arr_output)
    Arr_output(C, 1) = CDate(Arr_pitdata(C - 2, 1))
Next C
End Sub
Private Sub Get_CoreData_Daily_SupplierSensitive()
Dim Arr_Supplier()
Arr_Supplier = Get_SupplierList(Arr_pitdata)
Dim A%, B%, C%
For A = 1 To UBound(Arr_PNGroup)
    For B = 1 To UBound(Arr_SiteGroup)
        For C = LBound(Arr_Supplier) To UBound(Arr_Supplier)
            Call Get_Daily_GDSData_PNsite(Arr_PNGroup(A, 1), Arr_SiteGroup(B, 1), Arr_PNGroup(A, 2), Arr_SiteGroup(B, 2), CStr(Arr_Supplier(C)))
            If Need_Subtotal = True Then Call Get_Daily_GDSData_Subtotal(Arr_PNGroup(A, 1), Arr_SiteGroup(B, 1), CStr(Arr_Supplier(C))): Need_Subtotal = False
        Next C
    Next B
    If ExistWW = True And NeedWW = True Then Call Get_Daily_GDSData_WW(Arr_PNGroup(A, 1), Arr_PNGroup(A, 2)): NeedWW = False
Next A
End Sub

'FRED_WU1 2018/8/10
Option Explicit
Public Arr_wkBracket(), DicTAM As Object

Sub Main_GDS_Weekly_Format()
    Daily_Weekly = "Weekly"
    Application.ScreenUpdating = False
    Call Get_Foxjua_Inv
    Call collect_Checkbox_data
    Call Get_DicTAM
    Dim Class_PrepareData As Object
    Set Class_PrepareData = New Cls_PrepareData
    Class_PrepareData.Arr_Measure = Array("Released Forecast", "Backlog", "Future Backlog", "Net Supply", "Inventory On-Hand")
    Class_PrepareData.Arr_Measure_Min = Array("Released Forecast", "Backlog", "Future Backlog", "Net Supply", "Inventory On-Hand", "SLC On-Hand")

    Class_PrepareData.Main_Prepare_Data
    Arr_PNGroup = Class_PrepareData.Arr_PNGroup
    Arr_SiteGroup = Class_PrepareData.Arr_SiteGroup
    Arr_pitdata = Class_PrepareData.Arr_pitdata
    ExistWW = Class_PrepareData.ExistWW
    Arr_StartColumnDateWK = Class_PrepareData.Arr_StartColumnDateWK
    Set Class_PrepareData = Nothing
    If Supplier_Senstive = False Then
        Call Get_Dic_Index_Pitdata
    Else: Call Get_Dic_Index_Pitdata_SupplierSenstive
    End If

    Call Get_Output_FistRow
    Call Get_Arr_Formula_DSI
    Call Get_Arr_wkBracket

    If Supplier_Senstive = False Then
        Call Get_CoreData
    Else: Call Get_CoreData_SupplierSenstive
    End If
    If UBound(Arr_output, 2) = 1 Then MsgBox ("No data was presented ,please check whether Pitdata/Setting is correct!"): End
    Workbooks.Add
    Do While Sheets.Count > 1
        Sheets(Sheets.Count).Delete
    Loop
    ActiveSheet.Range(Cells(1, 1), Cells(UBound(Arr_output, 2), 1)).NumberFormat = "@"
    ActiveSheet.Name = "Weekly GDS"
    Cells(1, 1).Resize(UBound(Arr_output, 2), UBound(Arr_output)) = Application.Transpose(Arr_output)
    Dim Class_Format As Object
    Set Class_Format = New Cls_Format
    Call Class_Format.Formating_GDS

    Application.ScreenUpdating = True
    Call ClearRAM

End Sub

Private Sub Get_DicTAM()
Dim reg As Object, matchs As Object, A%, B%, C%, Dic_PNgroup As Object, T%
Dim Arr_Data_TAM()
Set reg = CreateObject("vbscript.regexp")
With Sheets("TAM")
    Arr_Data_TAM = .Range(.Cells(1, 1), .Cells(.[A65536].End(xlUp).Row, 4)).Value
End With
Set DicTAM = CreateObject("Scripting.dictionary")
With reg
    .Global = True
    .Pattern = "[A-Za-z0-9]{5}"
    For B = 1 To UBound(Arr_Data_TAM)
        If Len(Arr_Data_TAM(B, 1)) >= 5 Then
            ReDim Arr_temp(1 To 3)
            Arr_temp(1) = Arr_Data_TAM(B, 2) / 100
            Arr_temp(2) = Arr_Data_TAM(B, 3)
            Arr_temp(3) = Arr_Data_TAM(B, 4)

            Set matchs = .Execute(Arr_Data_TAM(B, 1))
            For A = 1 To matchs.Count
                If Not DicTAM.EXISTS(matchs.Item(A - 1).Value) Then
                    DicTAM(matchs.Item(A - 1).Value) = Arr_temp
                End If
            Next A
        End If
    Next B
End With
Erase Arr_Data_TAM
Set reg = Nothing: Set matchs = Nothing
End Sub
Private Sub Get_Arr_wkBracket()
Dim A&, B&
ReDim Arr_wkBracket(LBound(Arr_output) To UBound(Arr_output))
For A = LBound(Arr_output) To UBound(Arr_output)
    If IsDate(Arr_output(A, 1)) = True Then
        ReDim Arr_temp(1 To 2)
        For B = 9 To UBound(Arr_pitdata)
            If CDate(Arr_pitdata(B, 1)) = Arr_output(A, 1) Then
                Arr_temp(1) = B
                If Arr_temp(2) = Empty Then Arr_temp(2) = B
            End If
            If CDate(Arr_pitdata(B, 1)) = Arr_output(A, 1) + 6 Then Arr_temp(2) = B
        Next B
        Arr_wkBracket(A) = Arr_temp
    End If
Next A
End Sub
Private Sub Get_Output_FistRow()
Dim Arr_Title, A%, B%, C%
If Supplier_Senstive = False Then
Arr_Title = Array("Part-Group", "Description", "Site-Group", "CFG", "Buyer Name", "Measure", "BLG/INV", Arr_StartColumnDateWK(3))
Else: Arr_Title = Array("Part-Group", "Description", "Site-Group", "Supplier", "CFG", "Buyer Name", "Measure", "BLG/INV", Arr_StartColumnDateWK(3))
End If

Dim T%
ReDim Arr_Title_Sum(1 To 1)
For A = 1 To UBound(Arr_Title) + 1
    T = T + 1
    ReDim Preserve Arr_Title_Sum(1 To T)
    Arr_Title_Sum(T) = Arr_Title(A - 1)
Next A
For B = 1 To 2
    T = T + 1
    ReDim Preserve Arr_Title_Sum(1 To T)
    Arr_Title_Sum(T) = Arr_StartColumnDateWK(3) + B * 7
Next B
For C = 10 To UBound(Arr_pitdata)
    If CDate(Arr_pitdata(C, 1)) > Arr_StartColumnDateWK(3) + 3 * 7 - 1 Then
        T = T + 1
        ReDim Preserve Arr_Title_Sum(1 To T)
        Arr_Title_Sum(T) = CDate(Arr_pitdata(C, 1))
    End If
Next C

Dim D%
ReDim Arr_output(1 To UBound(Arr_Title_Sum), 1 To 1)
For D = 1 To UBound(Arr_Title_Sum)
    Arr_output(D, 1) = Arr_Title_Sum(D)
Next D

End Sub
Private Sub Get_Arr_Formula_DSI()
ReDim Arr_Formula_DSI(1 To UBound(Arr_output, 1))
Dim B%
For B = UBound(Arr_output, 1) To 10 Step -1
    If Arr_output(B, 1) - Arr_output(B - 1, 1) > 7 Then
        ReDim Arr_Formula_DSI(1 To B - 6)
    Else: Exit For
    End If
Next B

Dim A%
For A = 7 To UBound(Arr_Formula_DSI)
    If IsDate(Arr_output(A, 1)) = True Then
        Arr_Formula_DSI(A) = "=IFERROR(R[-1]C/(SUM(R[-4]C[1]:R[-4]C[4])/20),0)"
    End If
Next A
End Sub

'FRED_WU1 2018/8/10
Option Explicit
Dim Dic_PNs As Object
Public TAM_RollBack As Boolean
Public Dic_FSJ_Inv As Object
Public Function ClearRAM()
'DataFileLocation = Empty
Erase Arr_StartColumnDateWK, Arr_wkBracket, Arr_PNGroup, Arr_SiteGroup, Arr_pitdata, Arr_output, Arr_Formula_DSI
ExistWW = False
NeedWW = False
Lazy_Mode = False
TAM_RollBack = False
Supplier_Senstive = False
Set Dic_Index_Pitdata = Nothing
Daily_Weekly = Empty
Set Dic_FSJ_Inv = Nothing
End Function
Public Sub Get_Foxjua_Inv()
Dim Arr_INV_FSJ()
Arr_INV_FSJ = Sheets("WWT stock").UsedRange.Value
Set Dic_FSJ_Inv = CreateObject("scripting.dictionary")
Dim A&, Qty_temp&
For A = 2 To UBound(Arr_INV_FSJ)
    If Not Dic_FSJ_Inv.EXISTS(Arr_INV_FSJ(A, 1) & Arr_INV_FSJ(A, 2)) Then
        Dic_FSJ_Inv(Arr_INV_FSJ(A, 1) & Arr_INV_FSJ(A, 2)) = Arr_INV_FSJ(A, 3)
    Else
        Qty_temp = Dic_FSJ_Inv(Arr_INV_FSJ(A, 1) & Arr_INV_FSJ(A, 2)) + Arr_INV_FSJ(A, 3)
        Dic_FSJ_Inv(Arr_INV_FSJ(A, 1) & Arr_INV_FSJ(A, 2)) = Qty_temp
    End If
Next A
End Sub
Public Sub collect_Checkbox_data()
If ActiveSheet.CheckBoxes("Check Box 1").Value = xlOn Then Supplier_Senstive = True
If ActiveSheet.CheckBoxes("Check Box 2").Value = xlOn Then Lazy_Mode = True
If ActiveSheet.CheckBoxes("Check Box 3").Value = xlOn Then TAM_RollBack = True
If Lazy_Mode = False And ([A65536].End(xlUp).Row = 1 Or [C65536].End(xlUp).Row = 1) Then MsgBox ("Please input PN/Site or check LazyMode."): End
End Sub
Public Function Get_StartWKfromDate(InputData)
Dim X%, Y%
For X = 0 To 6
    If Weekday(InputData - X) = 7 Then Get_StartWKfromDate = InputData - X: Exit Function
Next X
If Get_StartWKfromDate = Empty Then MsgBox ("Start Date issue. please check pitdata."): End
End Function
Public Function Get_Column_StartSite(Column_Start, CheckPoint1, CheckPoint2)
If Arr_pitdata(Column_Start, CheckPoint1) <> Empty Or Arr_pitdata(Column_Start, CheckPoint2) <> Empty Then
    Get_Column_StartSite = Column_Start
    Else: Get_Column_StartSite = Column_Start - 1
End If
If Get_Column_StartSite <= 8 Then Get_Column_StartSite = 9
End Function
Public Function Get_SupplierList(Arr_pitdata)
Dim Dic_Single, A&, Arr_Single
ReDim Arr_Single(1 To 1)
Set Dic_Single = CreateObject("scripting.dictionary")
For A = 2 To UBound(Arr_pitdata, 2)
    If Not Dic_Single.EXISTS(Arr_pitdata(2, A)) Then
        Dic_Single.Add Arr_pitdata(2, A), ""
        Arr_Single(UBound(Arr_Single)) = Arr_pitdata(2, A)
        ReDim Preserve Arr_Single(1 To UBound(Arr_Single) + 1)
    End If
Next A
ReDim Preserve Arr_Single(1 To UBound(Arr_Single) - 1)
Get_SupplierList = Arr_Single
End Function
Public Sub Get_Dic_Index_Pitdata()
Set Dic_Index_Pitdata = CreateObject("scripting.dictionary")
Dim Arr_temp(), A&
For A = 2 To UBound(Arr_pitdata, 2)
    If Arr_pitdata(8, A) = "Released Forecast" Then
        If Not Dic_Index_Pitdata.EXISTS(Arr_pitdata(1, A) & Arr_pitdata(3, A)) Then
            Arr_temp = Array(A)
            Dic_Index_Pitdata(Arr_pitdata(1, A) & Arr_pitdata(3, A)) = Arr_temp
        Else
            Arr_temp = Dic_Index_Pitdata(Arr_pitdata(1, A) & Arr_pitdata(3, A))
            ReDim Preserve Arr_temp(LBound(Arr_temp) To UBound(Arr_temp) + 1)
            Arr_temp(UBound(Arr_temp)) = A
            Dic_Index_Pitdata(Arr_pitdata(1, A) & Arr_pitdata(3, A)) = Arr_temp
        End If
    End If
Next A
End Sub
Public Sub Get_Dic_Index_Pitdata_SupplierSenstive()
Set Dic_Index_Pitdata = CreateObject("scripting.dictionary")
Dim Arr_temp(), A&
For A = 2 To UBound(Arr_pitdata, 2)
    If Arr_pitdata(8, A) = "Released Forecast" Then
        If Not Dic_Index_Pitdata.EXISTS(Arr_pitdata(1, A) & Arr_pitdata(3, A) & Arr_pitdata(2, A)) Then
            Arr_temp = Array(A)
            Dic_Index_Pitdata(Arr_pitdata(1, A) & Arr_pitdata(3, A) & Arr_pitdata(2, A)) = Arr_temp
        Else
            Arr_temp = Dic_Index_Pitdata(Arr_pitdata(1, A) & Arr_pitdata(3, A) & Arr_pitdata(2, A))
            ReDim Preserve Arr_temp(LBound(Arr_temp) To UBound(Arr_temp) + 1)
            Arr_temp(UBound(Arr_temp)) = A
            Dic_Index_Pitdata(Arr_pitdata(1, A) & Arr_pitdata(3, A) & Arr_pitdata(2, A)) = Arr_temp
        End If
    End If
Next A
End Sub
Sub CheckBox0_Click()
If ActiveSheet.CheckBoxes("Check Box 0").Value = xlOn Then
    ActiveSheet.CheckBoxes("Check Box 0").Interior.ColorIndex = 33
    ActiveSheet.Shapes("TextBox 1").Fill.ForeColor.RGB = RGB(0, 176, 240)
Else
    ActiveSheet.CheckBoxes("Check Box 0").Interior.ColorIndex = 20
    ActiveSheet.Shapes("TextBox 1").Fill.ForeColor.RGB = RGB(255, 255, 255)
End If
End Sub
Sub CheckBox1_Click()
If ActiveSheet.CheckBoxes("Check Box 1").Value = xlOn Then
    ActiveSheet.CheckBoxes("Check Box 1").Interior.ColorIndex = 33
Else: ActiveSheet.CheckBoxes("Check Box 1").Interior.ColorIndex = 20
End If
End Sub
Sub CheckBox2_Click()
If ActiveSheet.CheckBoxes("Check Box 2").Value = xlOn Then
    ActiveSheet.CheckBoxes("Check Box 2").Interior.ColorIndex = 33
Else: ActiveSheet.CheckBoxes("Check Box 2").Interior.ColorIndex = 20
End If
End Sub
Sub CheckBox3_Click()
If ActiveSheet.CheckBoxes("Check Box 3").Value = xlOn Then
    ActiveSheet.CheckBoxes("Check Box 3").Interior.ColorIndex = 33
Else: ActiveSheet.CheckBoxes("Check Box 3").Interior.ColorIndex = 20
End If
End Sub
Sub Textbox1_Click()
Dim X$
X = Application.GetOpenFilename(, , "Please Select Pitdata")
If X <> "False" Then
ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Characters.TEXT = X
End If

End Sub
'FRED_WU1 2018/8/10
Option Explicit
Public Arr_Formula_DSI()
Public Supplier_Senstive As Boolean, Lazy_Mode As Boolean
Public Arr_PNGroup(), Arr_SiteGroup()
Public Arr_StartColumnDateWK()
Public Arr_pitdata(), Dic_Index_Pitdata As Object, Arr_output(), ExistWW As Boolean, NeedWW As Boolean
Public Daily_Weekly$
Public Sub Get_CoreData()
Dim A%, B%, Row_Temp&
For A = 1 To UBound(Arr_PNGroup)
    Row_Temp = UBound(Arr_output, 2) + 1
    NeedWW = False
    For B = 1 To UBound(Arr_SiteGroup)
        Call Get_GDSrow_PNsite(Arr_PNGroup(A, 1), Arr_SiteGroup(B, 1), Arr_PNGroup(A, 2), Arr_SiteGroup(B, 2))
    Next B
    If ExistWW = True And NeedWW = True Then Call Get_GDSrow_WW(Arr_PNGroup(A, 1), Row_Temp)
Next A
End Sub
Public Sub Get_CoreData_SupplierSenstive()
Dim A%, B%, C%, Arr_Supplier(), Row_Temp&
Arr_Supplier = Get_SupplierList(Arr_pitdata)
For A = 1 To UBound(Arr_PNGroup)
    Row_Temp = UBound(Arr_output, 2) + 1
    NeedWW = False
    For B = 1 To UBound(Arr_SiteGroup)
        For C = LBound(Arr_Supplier) To UBound(Arr_Supplier)
            Call Get_GDSrow_PNsiteSupplier(Arr_PNGroup(A, 1), Arr_SiteGroup(B, 1), Arr_PNGroup(A, 2), Arr_SiteGroup(B, 2), Arr_Supplier(C))
        Next C
    Next B
    If ExistWW = True And NeedWW = True Then Call Get_GDSrow_WW_SupplierSenstive(Arr_PNGroup(A, 1), Row_Temp)
Next A
End Sub
Private Sub Get_GDSrow_PNsite(PNgroup, Sitegroup, Arr_PNs, Arr_Sites)
Dim A%, B%, C%, T&, Added_NewRows As Boolean, Arr_temp_row()
Added_NewRows = False
For A = LBound(Arr_PNs) To UBound(Arr_PNs)
    For B = LBound(Arr_Sites) To UBound(Arr_Sites)
        If Dic_Index_Pitdata.EXISTS(Arr_PNs(A) & Arr_Sites(B)) Then
            NeedWW = True
            Arr_temp_row = Dic_Index_Pitdata(Arr_PNs(A) & Arr_Sites(B))
            If Added_NewRows = False Then
                Added_NewRows = True
                T = UBound(Arr_output, 2)
                ReDim Preserve Arr_output(1 To UBound(Arr_output), 1 To T + 5)
                For C = T + 1 To T + 5
                    Arr_output(1, C) = PNgroup                          'Part-Group
                    Arr_output(2, C) = Arr_pitdata(5, Arr_temp_row(0))      'Description
                    Arr_output(3, C) = Sitegroup                        'Site-Group
                    Arr_output(4, C) = Arr_pitdata(4, Arr_temp_row(0))      'CFG
                    Arr_output(5, C) = Arr_pitdata(6, Arr_temp_row(0))      'Buyer Name
                Next C
                Arr_output(6, T + 1) = "Released Forecast" 'Measure
                Arr_output(6, T + 2) = "Backlog"
                Arr_output(6, T + 3) = "Inventory/Net Supply"
                Arr_output(6, T + 4) = "Projected Inventory"
                Arr_output(6, T + 5) = "Projected DSI"
                Arr_output(7, T + 4) = "=R[-1]C-R[-2]C-R[-3]C"
            End If
            If Daily_Weekly = "Daily" Then
                Call Fill_CoreData_Integrated(T, 7, Arr_temp_row)
            Else
                Call Fill_CoreData_Weekly(T, 7, Arr_temp_row)
            End If
            Call Attach_FSJ_INV(T, 7, Arr_PNs(A), Arr_Sites(B))
        End If
    Next B
Next A
End Sub
Private Sub Get_GDSrow_PNsiteSupplier(PNgroup, Sitegroup, Arr_PNs, Arr_Sites, Supplier)
Dim Added_NewRows As Boolean
Added_NewRows = False
Dim A%, B%, C&, T&, Arr_temp_row()
For A = LBound(Arr_PNs) To UBound(Arr_PNs)
    For B = LBound(Arr_Sites) To UBound(Arr_Sites)
        If Dic_Index_Pitdata.EXISTS(Arr_PNs(A) & Arr_Sites(B) & Supplier) Then
            NeedWW = True
            Arr_temp_row = Dic_Index_Pitdata(Arr_PNs(A) & Arr_Sites(B) & Supplier)
            If Added_NewRows = False Then
                Added_NewRows = True
                T = UBound(Arr_output, 2)
                ReDim Preserve Arr_output(1 To UBound(Arr_output), 1 To T + 5)
                For C = T + 1 To T + 5
                    Arr_output(1, C) = PNgroup                          'Part-Group
                    Arr_output(2, C) = Arr_pitdata(5, Arr_temp_row(0))      'Description
                    Arr_output(3, C) = Sitegroup                        'Site-Group
                    Arr_output(4, C) = Supplier                           ' "Supplier"
                    Arr_output(5, C) = Arr_pitdata(4, Arr_temp_row(0))      'CFG
                    Arr_output(6, C) = Arr_pitdata(6, Arr_temp_row(0))      'Buyer Name
                Next C
                Arr_output(7, T + 1) = "Released Forecast"  'Measure
                Arr_output(7, T + 2) = "Backlog"
                Arr_output(7, T + 3) = "Inventory/Net Supply"
                Arr_output(7, T + 4) = "Projected Inventory"
                Arr_output(7, T + 5) = "Projected DSI"
                Arr_output(8, T + 4) = "=R[-1]C-R[-2]C-R[-3]C"
            End If
            If Daily_Weekly = "Daily" Then
                Call Fill_CoreData_Integrated(T, 8, Arr_temp_row)

            Else
                Call Fill_CoreData_Weekly(T, 8, Arr_temp_row)
            End If
            Call Attach_FSJ_INV(T, 8, Arr_PNs(A), Arr_Sites(B))
        End If
    Next B
Next A
End Sub
Private Sub Attach_FSJ_INV(T, Col_blg, PN, site)
If Dic_FSJ_Inv.EXISTS(PN & site) Then
    Arr_output(Col_blg, T + 3) = Arr_output(Col_blg, T + 3) + Dic_FSJ_Inv(PN & site)
End If
End Sub
Private Sub Get_GDSrow_WW(PNgroup, Row_Temp)
Dim T&, C&, D%, E%
T = UBound(Arr_output, 2)
ReDim Preserve Arr_output(1 To UBound(Arr_output), 1 To T + 5)
For C = T + 1 To T + 5
    Arr_output(1, C) = PNgroup               'Part-Group
    Arr_output(2, C) = Arr_output(2, T)      'Description
    Arr_output(3, C) = "WW"                  'Site-Group
    Arr_output(4, C) = Arr_output(4, T)      'CFG
    Arr_output(5, C) = Arr_output(5, T)      'Buyer Name
Next C
Arr_output(6, T + 1) = "Released Forecast"  'Measure
Arr_output(6, T + 2) = "Backlog"
Arr_output(6, T + 3) = "Inventory/Net Supply"
Arr_output(6, T + 4) = "Projected Inventory"
Arr_output(6, T + 5) = "Projected DSI"
Arr_output(7, T + 4) = "=sum(R[-1]C)-sum(R[-2]C)"
For D = 7 To UBound(Arr_output)
    Arr_output(D, T + 1) = "=IF(SUMIF(R" & Row_Temp & "C6:R[-1]C6,RC6,R" & Row_Temp & "C:R[-1]C)=0,"""",SUMIF(R" & Row_Temp & "C6:R[-1]C6,RC6,R" & Row_Temp & "C:R[-1]C))"
    Arr_output(D, T + 2) = "=IF(SUMIF(R" & Row_Temp & "C6:R[-2]C6,RC6,R" & Row_Temp & "C:R[-2]C)=0,"""",SUMIF(R" & Row_Temp & "C6:R[-2]C6,RC6,R" & Row_Temp & "C:R[-2]C))"
    Arr_output(D, T + 3) = "=IF(SUMIF(R" & Row_Temp & "C6:R[-3]C6,RC6,R" & Row_Temp & "C:R[-3]C)=0,"""",SUMIF(R" & Row_Temp & "C6:R[-3]C6,RC6,R" & Row_Temp & "C:R[-3]C))"
Next D
For E = 8 To UBound(Arr_output)
    Arr_output(E, T + 4) = "=SUM(R[-1]C7:R[-1]C)-SUM(R[-3]C7:R[-2]C)"
    If E < UBound(Arr_Formula_DSI) Then Arr_output(E, T + 5) = Arr_Formula_DSI(E)
Next E

End Sub
Private Sub Get_GDSrow_WW_SupplierSenstive(PNgroup, Row_Temp)
Dim T&, C&, D%, E%
T = UBound(Arr_output, 2)
ReDim Preserve Arr_output(1 To UBound(Arr_output), 1 To T + 5)
For C = T + 1 To T + 5
    Arr_output(1, C) = PNgroup               'Part-Group
    Arr_output(2, C) = Arr_output(2, T)      'Description
    Arr_output(3, C) = "WW"                  'Site-Group
    Arr_output(4, C) = Arr_output(4, T)
    Arr_output(5, C) = Arr_output(5, T)      'CFG
    Arr_output(6, C) = Arr_output(6, T)      'Buyer Name
Next C
Arr_output(7, T + 1) = "Released Forecast"  'Measure
Arr_output(7, T + 2) = "Backlog"
Arr_output(7, T + 3) = "Inventory/Net Supply"
Arr_output(7, T + 4) = "Projected Inventory"
Arr_output(7, T + 5) = "Projected DSI"
Arr_output(8, T + 4) = "=SUM(R[-1]C)-SUM(R[-2]C)"
For D = 8 To UBound(Arr_output)
    Arr_output(D, T + 1) = "=IF(SUMIF(R" & Row_Temp & "C7:R[-1]C7,RC7,R" & Row_Temp & "C:R[-1]C)=0,"""",SUMIF(R" & Row_Temp & "C7:R[-1]C7,RC7,R" & Row_Temp & "C:R[-1]C))"
    Arr_output(D, T + 2) = "=IF(SUMIF(R" & Row_Temp & "C7:R[-2]C7,RC7,R" & Row_Temp & "C:R[-2]C)=0,"""",SUMIF(R" & Row_Temp & "C7:R[-2]C7,RC7,R" & Row_Temp & "C:R[-2]C))"
    Arr_output(D, T + 3) = "=IF(SUMIF(R" & Row_Temp & "C7:R[-3]C7,RC7,R" & Row_Temp & "C:R[-3]C)=0,"""",SUMIF(R" & Row_Temp & "C7:R[-3]C7,RC7,R" & Row_Temp & "C:R[-3]C))"
Next D

For E = 9 To UBound(Arr_output)
    Arr_output(E, T + 4) = "=SUM(R[-1]C8:R[-1]C)-SUM(R[-3]C8:R[-2]C)"
    If E < UBound(Arr_Formula_DSI) Then Arr_output(E, T + 5) = Arr_Formula_DSI(E)
Next E
End Sub
Private Sub Fill_CoreData_Integrated(T, Col_blg, Arr_temp_row)
Dim D&, E%, G%, StartColumn_Temp%, Arr_Temp_Col(), gap_Pit_out%
For D = LBound(Arr_temp_row) To UBound(Arr_temp_row)
    StartColumn_Temp = Get_Column_StartSite(Arr_StartColumnDateWK(1), Arr_temp_row(D) + 1, Arr_temp_row(D) + 4)
    If Arr_pitdata(StartColumn_Temp, Arr_temp_row(D) + 1) <> Empty Then
        Arr_output(Col_blg, T + 2) = Arr_output(Col_blg, T + 2) + CLng(Arr_pitdata(StartColumn_Temp, Arr_temp_row(D) + 1))
    End If
    If Arr_pitdata(StartColumn_Temp, Arr_temp_row(D) + 4) <> Empty Then
        Arr_output(Col_blg, T + 3) = Arr_output(Col_blg, T + 3) + CLng(Arr_pitdata(StartColumn_Temp, Arr_temp_row(D) + 4))
    End If

    gap_Pit_out = UBound(Arr_pitdata) - UBound(Arr_output)
    For E = Col_blg + 1 To UBound(Arr_output)
        If E + gap_Pit_out >= StartColumn_Temp Then
            If Arr_pitdata(E + gap_Pit_out, Arr_temp_row(D)) <> Empty Then
                Arr_output(E, T + 1) = Arr_output(E, T + 1) + CLng(Arr_pitdata(E + gap_Pit_out, Arr_temp_row(D)))
            End If
            If Arr_pitdata(E + gap_Pit_out, Arr_temp_row(D) + 2) <> Empty Then
                Arr_output(E, T + 2) = Arr_output(E, T + 2) + CLng(Arr_pitdata(E + gap_Pit_out, Arr_temp_row(D) + 2))
            End If
            If Arr_pitdata(E + gap_Pit_out, Arr_temp_row(D) + 3) <> Empty Then
                Arr_output(E, T + 3) = Arr_output(E, T + 3) + CLng(Arr_pitdata(E + gap_Pit_out, Arr_temp_row(D) + 3))
            End If
        End If
    Next E
Next D
For G = Col_blg + 1 To UBound(Arr_output)
    Arr_output(G, T + 4) = "=SUM(R[-1]C" & Col_blg & ":R[-1]C)-SUM(R[-3]C" & Col_blg & ":R[-2]C)"
    If G < UBound(Arr_output) - 3 Then
        Arr_output(G, T + 5) = Arr_Formula_DSI(G)
    End If
Next G
End Sub
Private Sub Fill_CoreData_Weekly(T, Col_blg, Arr_temp_row)
Dim D&, E%, F&, G%, H%, StartColumn_Temp%, Arr_Temp_Col(), Arr_temp()
For D = LBound(Arr_temp_row) To UBound(Arr_temp_row)
    StartColumn_Temp = Get_Column_StartSite(Arr_StartColumnDateWK(1), Arr_temp_row(D) + 1, Arr_temp_row(D) + 4)
    If Arr_pitdata(StartColumn_Temp, Arr_temp_row(D) + 1) <> Empty Then
        Arr_output(Col_blg, T + 2) = Arr_output(Col_blg, T + 2) + CLng(Arr_pitdata(StartColumn_Temp, Arr_temp_row(D) + 1))
    End If
    If Arr_pitdata(StartColumn_Temp, Arr_temp_row(D) + 4) <> Empty Then
        Arr_output(Col_blg, T + 3) = Arr_output(Col_blg, T + 3) + CLng(Arr_pitdata(StartColumn_Temp, Arr_temp_row(D) + 4))
    End If
    For E = Col_blg + 1 To UBound(Arr_output)
        Arr_Temp_Col = Arr_wkBracket(E)
        If E = Col_blg + 1 Then Arr_Temp_Col(1) = StartColumn_Temp
        For F = Arr_Temp_Col(1) To Arr_Temp_Col(2)
'            If Arr_pitdata(F, Arr_temp_row(D)) <> Empty Then
'                Arr_output(E, T + 1) = Arr_output(E, T + 1) + CLng(Arr_pitdata(F, Arr_temp_row(D)))
'            End If
            If Arr_pitdata(F, Arr_temp_row(D) + 2) <> Empty Then
                Arr_output(E, T + 2) = Arr_output(E, T + 2) + CLng(Arr_pitdata(F, Arr_temp_row(D) + 2))
            End If
            If Arr_pitdata(F, Arr_temp_row(D) + 3) <> Empty Then
                Arr_output(E, T + 3) = Arr_output(E, T + 3) + CLng(Arr_pitdata(F, Arr_temp_row(D) + 3))
            End If
        Next F
        If TAM_RollBack = False Or Not DicTAM.EXISTS(Arr_pitdata(1, Arr_temp_row(D))) Then
            For H = Arr_Temp_Col(1) To Arr_Temp_Col(2)
                If Arr_pitdata(H, Arr_temp_row(D)) <> Empty Then
                    Arr_output(E, T + 1) = Arr_output(E, T + 1) + CLng(Arr_pitdata(H, Arr_temp_row(D)))
                End If
            Next H
        Else
            Arr_temp = DicTAM(Arr_pitdata(1, Arr_temp_row(D)))
            For H = Arr_Temp_Col(1) To Arr_Temp_Col(2)
                If Arr_pitdata(H, Arr_temp_row(D)) <> Empty Then
                    If Arr_output(E, 1) >= Arr_temp(2) And Arr_output(E, 1) <= Arr_temp(3) Then
                        Arr_output(E, T + 1) = Arr_output(E, T + 1) + CLng(Arr_pitdata(H, Arr_temp_row(D))) / Arr_temp(1)
                    Else
                        Arr_output(E, T + 1) = Arr_output(E, T + 1) + CLng(Arr_pitdata(H, Arr_temp_row(D)))
                    End If
                End If
            Next H
        End If
    Next E
Next D
For G = Col_blg + 1 To UBound(Arr_output)
    Arr_output(G, T + 4) = "=SUM(R[-1]C" & Col_blg & ":R[-1]C)-SUM(R[-3]C" & Col_blg & ":R[-2]C)"
    If G < UBound(Arr_Formula_DSI) Then Arr_output(G, T + 5) = Arr_Formula_DSI(G)
Next G
End Sub

'FRED_WU1 2018/8/10
Option Explicit
Public Sub Formating_GDS()
Dim A&, B%, Arr_Widths, Column_Measurement%
If Supplier_Senstive = False Then
    Column_Measurement = 6
    Arr_Widths = Array(10, 5, 10, 5, 12, 19, 8)
Else
    Column_Measurement = 7
    Arr_Widths = Array(10, 5, 10, 12, 5, 12, 19, 8)
End If
Range(Cells(1, 1), Cells(1, Column_Measurement + 1)).Interior.ColorIndex = 6
With Range(Cells(1, Column_Measurement + 2), Cells(1, UBound(Arr_output, 1)))
    .Interior.ColorIndex = 20
    .NumberFormat = "m/d;@"
End With
Range(Cells(5, Column_Measurement + 1), Cells(5, UBound(Arr_output, 1))).NumberFormat = "0;[Red]-0"
Range(Cells(6, Column_Measurement + 2), Cells(6, UBound(Arr_output, 1))).NumberFormat = "0"
Call Add_ConditionFormat(Range(Cells(6, Column_Measurement + 2), Cells(6, UBound(Arr_output, 1))))
If UBound(Arr_output, 2) > 6 Then
    Range(Cells(2, 1), Cells(6, UBound(Arr_output, 1))).Copy
    Range(Cells(7, 1), Cells(UBound(Arr_output, 2), UBound(Arr_output, 1))).PasteSpecial (xlPasteFormats)
End If

Dim Thinline_rng As Range, Thickline_Rng As Range
A = 1
Do While A < UBound(Arr_output, 2)
    If Arr_output(1, A) <> Arr_output(1, A + 1) Then
        If Thickline_Rng Is Nothing Then
            Set Thickline_Rng = Range(Cells(A, 1), Cells(A, UBound(Arr_output, 1)))
        Else: Set Thickline_Rng = Union(Thickline_Rng, Range(Cells(A, 1), Cells(A, UBound(Arr_output, 1))))
        End If
    ElseIf Arr_output(3, A) <> Arr_output(3, A + 1) Or Arr_output(4, A) <> Arr_output(4, A + 1) Then
        If Thinline_rng Is Nothing Then
            Set Thinline_rng = Range(Cells(A, 1), Cells(A, UBound(Arr_output, 1)))
        Else: Set Thinline_rng = Union(Thinline_rng, Range(Cells(A, 1), Cells(A, UBound(Arr_output, 1))))
        End If
    End If
A = A + 5
Loop

Dim Rng_Monthly As Range
Set Rng_Monthly = Get_Monthly_Rng()
If Not Rng_Monthly Is Nothing Then Rng_Monthly.Interior.ColorIndex = 33

If Not Thinline_rng Is Nothing Then Thinline_rng.Borders(xlEdgeBottom).Weight = xlThin
If Not Thickline_Rng Is Nothing Then Thickline_Rng.Borders(xlEdgeBottom).Weight = xlThick
Range(Cells(UBound(Arr_output, 2), 1), Cells(UBound(Arr_output, 2), UBound(Arr_output, 1))).Borders(xlEdgeBottom).Weight = xlThick
For B = LBound(Arr_Widths) To UBound(Arr_Widths)
    Columns(B + 1).ColumnWidth = Arr_Widths(B)
Next B
Range(Cells(1, Column_Measurement + 2), Cells(1, UBound(Arr_output, 1))).ColumnWidth = 6
Range(Cells(1, Column_Measurement + 2), Cells(1, UBound(Arr_output, 1))).HorizontalAlignment = xlLeft
Cells(2, Column_Measurement + 2).Select
ActiveWindow.FreezePanes = True
Application.ErrorCheckingOptions.OmittedCells = False
End Sub
Public Sub Formating_GDS_DailyIntegrated()
Dim Rng_Wk1 As Range, Rng_Wk2 As Range, Rng_Wk3 As Range
Dim Rng_Wk1_G As Range, Rng_Wk2_G As Range, Rng_Wk3_G As Range

Dim A&
For A = 7 To UBound(Arr_output)
    If IsDate(Arr_output(A, 1)) Then
        If Arr_output(A, 1) >= Arr_StartColumnDateWK(3) And Arr_output(A, 1) < Arr_StartColumnDateWK(3) + 7 Then
            If Not Rng_Wk1 Is Nothing Then
                Set Rng_Wk1 = Union(Rng_Wk1, Cells(1, A))
            Else: Set Rng_Wk1 = Cells(1, A)
            End If
        ElseIf Arr_output(A, 1) >= Arr_StartColumnDateWK(3) + 7 And Arr_output(A, 1) < Arr_StartColumnDateWK(3) + 14 Then
            If Not Rng_Wk2 Is Nothing Then
                Set Rng_Wk2 = Union(Rng_Wk2, Cells(1, A))
            Else: Set Rng_Wk2 = Cells(1, A)
            End If
        ElseIf Arr_output(A, 1) >= Arr_StartColumnDateWK(3) + 14 And Arr_output(A, 1) < Arr_StartColumnDateWK(3) + 21 Then
            If Not Rng_Wk3 Is Nothing Then
                Set Rng_Wk3 = Union(Rng_Wk3, Cells(1, A))
            Else: Set Rng_Wk3 = Cells(1, A)
            End If
        End If


        If Arr_output(A, 1) >= Arr_StartColumnDateWK(3) And Arr_output(A, 1) < Arr_StartColumnDateWK(3) + 6 Then
            If Not Rng_Wk1_G Is Nothing Then
                Set Rng_Wk1_G = Union(Rng_Wk1_G, Cells(1, A))
            Else: Set Rng_Wk1_G = Cells(1, A)
            End If
        ElseIf Arr_output(A, 1) >= Arr_StartColumnDateWK(3) + 7 And Arr_output(A, 1) < Arr_StartColumnDateWK(3) + 13 Then
            If Not Rng_Wk2_G Is Nothing Then
                Set Rng_Wk2_G = Union(Rng_Wk2_G, Cells(1, A))
            Else: Set Rng_Wk2_G = Cells(1, A)
            End If
        ElseIf Arr_output(A, 1) >= Arr_StartColumnDateWK(3) + 14 And Arr_output(A, 1) < Arr_StartColumnDateWK(3) + 20 Then
            If Not Rng_Wk3_G Is Nothing Then
                Set Rng_Wk3_G = Union(Rng_Wk3_G, Cells(1, A))
            Else: Set Rng_Wk3_G = Cells(1, A)
            End If
        End If
    End If
Next A
Dim Rng_Monthly As Range
Set Rng_Monthly = Get_Monthly_Rng()
If Not Rng_Monthly Is Nothing Then Rng_Monthly.Interior.ColorIndex = 33

If Not Rng_Wk1 Is Nothing Then Rng_Wk1.Interior.ColorIndex = 43
If Not Rng_Wk2 Is Nothing Then Rng_Wk2.Interior.ColorIndex = 40
If Not Rng_Wk3 Is Nothing Then Rng_Wk3.Interior.ColorIndex = 37

Rng_Wk1.Resize(UBound(Arr_output, 2), Rng_Wk1.Cells.Count).Borders(xlEdgeRight).Weight = xlThin
Rng_Wk2.Resize(UBound(Arr_output, 2), Rng_Wk2.Cells.Count).Borders(xlEdgeRight).Weight = xlThin
Rng_Wk3.Resize(UBound(Arr_output, 2), Rng_Wk3.Cells.Count).Borders(xlEdgeRight).Weight = xlThin
Rng_Wk1_G.Columns.Group
Rng_Wk2_G.Columns.Group
Rng_Wk3_G.Columns.Group

End Sub
Private Sub Add_ConditionFormat(ByVal ConditionRng)
With ConditionRng.FormatConditions.Add(xlCellValue, xlEqual, 0)
    .Interior.ColorIndex = 4
End With
With ConditionRng.FormatConditions.Add(xlCellValue, xlLess, 10)
    .Interior.ColorIndex = 3
    .Font.ColorIndex = 2
End With
With ConditionRng.FormatConditions.Add(xlCellValue, xlBetween, 10, 17)
    .Interior.ColorIndex = 6
End With
With ConditionRng.FormatConditions.Add(xlCellValue, xlBetween, 17, 30)
    .Interior.ColorIndex = 4
End With
With ConditionRng.FormatConditions.Add(xlCellValue, xlGreaterEqual, 30)
    .Interior.ColorIndex = 5
    .Font.ColorIndex = 2
End With
Set ConditionRng = Nothing
End Sub
Public Sub Formating_GDS_Daily(Criterion_Date)
Dim A%, B&, C&, D%, DSI_rng As Range, INV_rng As Range, Thinline_rng As Range, Thickline_Rng As Range, xlMedium_Rng As Range
Dim LastCol%
LastCol = UBound(Arr_output, 1)
Range(Cells(1, 1), Cells(1, 10)).Interior.ColorIndex = 6
Range(Cells(1, 11), Cells(1, 11).End(xlToRight)).NumberFormat = "m/d;@"
Range(Cells(1, 11), Cells(1, 11).End(xlToRight)).HorizontalAlignment = xlLeft
For A = 11 To UBound(Arr_output, 1)
    If Arr_output(A + 1, 1) - Arr_output(A, 1) = 1 Then
    Cells(1, A).Interior.ColorIndex = 6
    Else: Range(Cells(1, A), Cells(1, LastCol)).Interior.ColorIndex = 20: Exit For
    End If
Next
For B = 2 To UBound(Arr_output, 2)
    If Arr_output(10, B) = "Projected Inventory" Or Arr_output(10, B) = "Subtotal Inventory" Or Arr_output(10, B) = "WW Total Inventory" Then
        If INV_rng Is Nothing Then
            Set INV_rng = Range(Cells(B, 11), Cells(B, LastCol))
        Else: Set INV_rng = Union(INV_rng, Range(Cells(B, 11), Cells(B, LastCol)))
        End If
    End If
    If Arr_output(10, B) = "Projected DSI" Or Arr_output(10, B) = "Subtotal DSI" Or Arr_output(10, B) = "WW Total DSI" Then
        If DSI_rng Is Nothing Then
            Set DSI_rng = Range(Cells(B, 11), Cells(B, LastCol))
        Else: Set DSI_rng = Union(DSI_rng, Range(Cells(B, 11), Cells(B, LastCol)))
        End If
    End If

    If Arr_output(10, B) = "Released Forecast" Then
        If Thinline_rng Is Nothing Then
            Set Thinline_rng = Range(Cells(B, 1), Cells(B, LastCol))
        Else: Set Thinline_rng = Union(Thinline_rng, Range(Cells(B, 1), Cells(B, LastCol)))
        End If
    End If

    If Arr_output(10, B) = "Subtotal Forecast" Then
        If xlMedium_Rng Is Nothing Then
            Set xlMedium_Rng = Range(Cells(B, 1), Cells(B + 4, LastCol))
        Else: Set xlMedium_Rng = Union(xlMedium_Rng, Range(Cells(B, 1), Cells(B + 4, LastCol)))
        End If
    End If

    If Arr_output(10, B) = "WW Total Forecast" Then
        If Thickline_Rng Is Nothing Then
            Set Thickline_Rng = Range(Cells(B, 1), Cells(B + 4, LastCol))
        Else: Set Thickline_Rng = Union(Thickline_Rng, Range(Cells(B, 1), Cells(B + 4, LastCol)))
        End If
    End If

Next B
Range(Cells(2, 1), Cells(UBound(Arr_output, 2), LastCol)).Interior.ColorIndex = 2

Dim Rng_Monthly As Range
Set Rng_Monthly = Get_Monthly_Rng()
If Not Rng_Monthly Is Nothing Then Rng_Monthly.Interior.ColorIndex = 33

If Not DSI_rng Is Nothing Then Call Add_ConditionFormat(DSI_rng)
DSI_rng.NumberFormat = "0": DSI_rng.HorizontalAlignment = xlCenter
If Not INV_rng Is Nothing Then INV_rng.NumberFormat = "0;[Red]-0"

If Not Thinline_rng Is Nothing Then Thinline_rng.Borders(xlEdgeTop).Weight = xlThin

If Not xlMedium_Rng Is Nothing Then
    xlMedium_Rng.Borders(xlEdgeTop).Weight = xlMedium
    xlMedium_Rng.Borders(xlEdgeBottom).Weight = xlMedium
End If

If Not Thickline_Rng Is Nothing Then
    Thickline_Rng.Borders(xlEdgeTop).Weight = xlThick
    Thickline_Rng.Borders(xlEdgeBottom).Weight = xlThick
End If

C = 2
Do While C < UBound(Arr_output, 2)
    Range(Cells(C, 1), Cells(Cells(C, 1).End(xlDown).Row, 1)).Rows.Group
    C = Cells(Cells(C, 1).End(xlDown).Row, 1).End(xlDown).Row
Loop
For D = 11 To UBound(Arr_output, 1)
    If Cells(1, D) = Criterion_Date Then
        If D > 13 Then Range(Cells(1, 11), Cells(1, D - 2)).EntireColumn.Group
        Range(Cells(1, D + 1), Cells([B65536].End(xlUp).Row, D + 1)).Borders(xlEdgeLeft).Weight = xlMedium: Exit For
    End If
Next
Dim Arr_Widths, E%
Arr_Widths = Array(5, 8, 5, 5, 7, 1, 1, 1, 1, 17)
For E = LBound(Arr_Widths) To UBound(Arr_Widths)
    Columns(E + 1).ColumnWidth = Arr_Widths(E)
Next E
Range(Cells(1, 11), Cells(1, UBound(Arr_output, 1))).ColumnWidth = 6

Range("K2").Select: ActiveWindow.FreezePanes = True
ActiveSheet.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1

End Sub
Private Function Get_Monthly_Rng()
Dim B%, Rng_Monthly As Range
For B = UBound(Arr_output) To 10 Step -1
    If IsDate(Arr_output(B, 1)) Then
        If Arr_output(B, 1) - Arr_output(B - 1, 1) > 7 Then
            If Not Rng_Monthly Is Nothing Then
                Set Rng_Monthly = Union(Rng_Monthly, Cells(1, B)): Set Rng_Monthly = Union(Rng_Monthly, Cells(1, B - 1))
            Else: Set Rng_Monthly = Union(Cells(1, B), Cells(1, B - 1))
            End If
        Else
        Exit For
        End If
    End If
Next B
Set Get_Monthly_Rng = Rng_Monthly
End Function

'FRED_WU1 2018/8/10
Option Explicit
Public ExistWW As Boolean
Public Arr_Measure, Arr_Measure_Min, Arr_PNGroup, Arr_SiteGroup, Arr_pitdata
Dim Dic_PNs As Object
Public Arr_StartColumnDateWK
Public Sub Main_Prepare_Data()
Dim CellValue$, DataFileLocation$
If ActiveSheet.CheckBoxes("Check Box 0").Value = xlOn Then CellValue = ActiveSheet.Shapes("TextBox 1").TextFrame2.TextRange.Characters.TEXT
If ExistsFile_UseFso(CellValue) Then
DataFileLocation = CellValue
ElseIf ExistsFile_UseFso(CellValue & ".txt") Then
DataFileLocation = CellValue & ".TXT"
ElseIf ExistsFile_UseFso(CellValue & ".xls") Then
DataFileLocation = CellValue & ".xls"
ElseIf ExistsFile_UseFso(CellValue & ".xlsx") Then
DataFileLocation = CellValue & ".xlsx"
Else
DataFileLocation = Application.GetOpenFilename(, , "Incorrect file location in Textbox1,Please Select PitData")
End If

If DataFileLocation = "False" Or DataFileLocation = Empty Then End
If UCase(DataFileLocation) Like "*.TXT" Then
    Arr_pitdata = Get_Arr_Data_FromTxt(DataFileLocation)
Else: Arr_pitdata = Get_Arr_Data_FromExcel(DataFileLocation)
End If
If IsArray(Arr_pitdata) = False Then MsgBox ("pitdata incorrect"): End

If Lazy_Mode = True Then
    Call Get_LazyModePNsite
Else
    Dim Str_temp$
    Arr_PNGroup = Application.Transpose(Application.index(Cells(2, 1).Resize([A65536].End(xlUp).Row - 1, 1), , 1))
    If IsArray(Arr_PNGroup) = False Then
        Str_temp = Arr_PNGroup
        ReDim Arr_PNGroup(1 To 1)
        Arr_PNGroup(1) = Str_temp
    End If
    Arr_SiteGroup = Application.Transpose(Application.index(Cells(2, 3).Resize([C65536].End(xlUp).Row - 1, 1), , 1))
    If IsArray(Arr_SiteGroup) = False Then
        Str_temp = Arr_SiteGroup
        ReDim Arr_SiteGroup(1 To 1)
        Arr_SiteGroup(1) = Str_temp
    End If
End If
Call Get_PNs_FromPNgroup
Call Get_Sites_FromSiteGroup
Call Get_Arr_StartColumnDateWK
Arr_pitdata = Condense_Arr_Pitdata()
End Sub
Private Sub Get_LazyModePNsite()
Dim Dic_PN As Object, Dic_CFG As Object, Dic_Site As Object
Set Dic_PN = CreateObject("scripting.dictionary")
Set Dic_CFG = CreateObject("scripting.dictionary")
Set Dic_Site = CreateObject("scripting.dictionary")
Dim A&, Str_temp$
For A = 2 To UBound(Arr_pitdata)
    If Not Dic_PN.EXISTS(Arr_pitdata(A, 1)) Then
        Dic_PN(Arr_pitdata(A, 1)) = ""
        If Not Dic_CFG.EXISTS(Arr_pitdata(A, 4)) Then
            Dic_CFG(Arr_pitdata(A, 4)) = Arr_pitdata(A, 1)
        Else
            Str_temp = Dic_CFG(Arr_pitdata(A, 4))
            Str_temp = Str_temp & ";" & Arr_pitdata(A, 1)
            Dic_CFG(Arr_pitdata(A, 4)) = Str_temp
        End If
    End If
    If Not Dic_Site.EXISTS(Arr_pitdata(A, 3)) Then
        Dic_Site(Arr_pitdata(A, 3)) = ""
    End If
Next A
Dim Arr_cfg(), Arr_site()
Arr_cfg = Dic_CFG.ITEMS
Arr_site = Dic_Site.KEYS
ReDim Arr_PNGroup(1 To Dic_CFG.Count): ReDim Arr_SiteGroup(1 To Dic_Site.Count + 1)
Dim B%, C%
For B = 1 To Dic_CFG.Count
    Arr_PNGroup(B) = Arr_cfg(B - 1)
Next B
For C = 1 To Dic_Site.Count
    Arr_SiteGroup(C) = Arr_site(C - 1)
Next C
Arr_SiteGroup(UBound(Arr_SiteGroup)) = "WW"
End Sub
Private Function Condense_Arr_Pitdata()
Dim Dic_Measure As Object
Dim A%, B%, C&, D%, T&
Dim E%, Dic_Measure_Min
Set Dic_Measure = CreateObject("SCRIPTING.DICTIONARY")
For A = LBound(Arr_Measure) To UBound(Arr_Measure)
    If Not Dic_Measure.EXISTS(Arr_Measure(A)) Then Dic_Measure(Arr_Measure(A)) = ""
Next A
Set Dic_Measure_Min = CreateObject("SCRIPTING.DICTIONARY")
For E = LBound(Arr_Measure_Min) To UBound(Arr_Measure_Min)
    If Not Dic_Measure_Min.EXISTS(Arr_Measure_Min(E)) Then Dic_Measure_Min(Arr_Measure_Min(E)) = ""
Next E


ReDim Arr_output(1 To UBound(Arr_pitdata, 2), 1 To 1)
For B = LBound(Arr_pitdata, 2) To UBound(Arr_pitdata, 2)
    Arr_output(B, 1) = Arr_pitdata(1, B)
Next B
T = 1
For C = 2 To UBound(Arr_pitdata)
    If Dic_Measure.EXISTS(Arr_pitdata(C, 8)) And Dic_PNs.EXISTS(Arr_pitdata(C, 1)) Then
        T = T + 1
        ReDim Preserve Arr_output(1 To UBound(Arr_output), 1 To T)
        For D = 1 To UBound(Arr_pitdata, 2)
            Arr_output(D, T) = Arr_pitdata(C, D)
        Next D
    End If
    If Dic_Measure_Min.EXISTS(Arr_pitdata(C, 8)) Then
        Dic_Measure_Min(Arr_pitdata(C, 8)) = "Y"
    End If
Next C
If UBound(Arr_output, 2) = 1 Then MsgBox ("No data was presented ,please check whether Pitdata/Setting is correct!"): End
Condense_Arr_Pitdata = Arr_output

Set Dic_Measure = Nothing
Dim F%, Arr_item()
Arr_item = Dic_Measure_Min.ITEMS
For F = LBound(Arr_item) To UBound(Arr_item)
    If Arr_item(F) = Empty Then
        MsgBox (" lack of required measurement in pitdata,required measurement are " & Chr(10) & Join(Arr_Measure_Min, ","))
        End
    End If
Next F
Erase Arr_output, Arr_pitdata, Arr_Measure, Arr_Measure_Min
Set Dic_Measure_Min = Nothing
End Function
Private Sub Get_PNs_FromPNgroup()
ReDim Arr_output(1 To 2, 1 To 1)
Dim reg As Object, matchs As Object, A%, B%, C%, Dic_PNgroup As Object, T%
Set reg = CreateObject("vbscript.regexp")

Set Dic_PNs = CreateObject("Scripting.dictionary")
Set Dic_PNgroup = CreateObject("Scripting.dictionary")
With reg
    .Global = True
    '.Pattern = "[A-Za-z0-9]{5}"
    .Pattern = "[A-Za-z0-9_\-]{4,}"
    For B = 1 To UBound(Arr_PNGroup)
        If Not Dic_PNgroup.EXISTS(Arr_PNGroup(B)) Then
            Dic_PNgroup(Arr_PNGroup(B)) = ""
            ReDim Arr_temp(1 To 1)
            Set matchs = .Execute(Arr_PNGroup(B))
            C = 0
            For A = 1 To matchs.Count
                C = C + 1
                ReDim Preserve Arr_temp(1 To C)
                Arr_temp(C) = UCase(matchs.Item(A - 1).Value)
                If Not Dic_PNs.EXISTS(Arr_temp(C)) Then Dic_PNs(Arr_temp(C)) = ""
            Next A
            T = T + 1
            ReDim Preserve Arr_output(1 To 2, 1 To T)
            Arr_output(1, T) = Arr_PNGroup(B)
            Arr_output(2, T) = Arr_temp
        End If
    Next B
End With
Arr_PNGroup = Transpose_Arr(Arr_output)
Erase Arr_output
Set reg = Nothing: Set matchs = Nothing
End Sub
Private Sub Get_Sites_FromSiteGroup()
ExistWW = False
Dim reg As Object, matchs As Object, Dic_Sitegroup As Object
Dim A%, B%, C%, T%
ReDim Arr_output(1 To 2, 1 To 1)
Set reg = CreateObject("vbscript.regexp")
Set Dic_Sitegroup = CreateObject("Scripting.dictionary")

With reg
    .Global = True
    .Pattern = "[A-Za-z0-9_\-]{1,}"
    For B = 1 To UBound(Arr_SiteGroup)
        If Not Dic_Sitegroup.EXISTS(Arr_SiteGroup(B)) Then
            Dic_Sitegroup(Arr_SiteGroup(B)) = ""
            If UCase(Arr_SiteGroup(B)) = "WW" Then ExistWW = True
            ReDim Arr_temp(1 To 1)
            Set matchs = .Execute(Arr_SiteGroup(B))
            C = 0
            For A = 1 To matchs.Count
                C = C + 1
                ReDim Preserve Arr_temp(1 To C)
                Arr_temp(C) = matchs.Item(A - 1).Value
            Next A
            T = T + 1
            ReDim Preserve Arr_output(1 To 2, 1 To T)
            Arr_output(1, T) = Arr_SiteGroup(B)
            Arr_output(2, T) = Arr_temp
        End If
    Next B
End With
Arr_SiteGroup = Transpose_Arr(Arr_output)
Set reg = Nothing: Set matchs = Nothing
End Sub
Private Function Get_Arr_Data_FromExcel(File_Location As String)
Dim Arr_Data()
Workbooks.Open File_Location
If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData
Arr_Data = ActiveSheet.UsedRange.Value
ActiveWorkbook.Close False
Get_Arr_Data_FromExcel = Arr_Data
Erase Arr_Data
End Function
Private Function Get_Arr_Data_FromTxt(File_Location)
Dim Arr_rows
Dim Total_Row&, Total_Col%
Dim A&, B%, Arr_temp
Open File_Location For Input As #1
Arr_rows = Split(StrConv(InputB(LOF(1), 1), vbUnicode), Chr(10))
Close #1
Total_Col = UBound(Split(Arr_rows(0), Chr(9))) + 1
Total_Row = UBound(Arr_rows) + 1
ReDim Arr_Data(1 To Total_Row - 1, 1 To Total_Col)
For A = 1 To Total_Row - 1
    Arr_temp = Split(Arr_rows(A - 1), Chr(9))
    For B = 1 To Total_Col
        Arr_Data(A, B) = Arr_temp(B - 1)
    Next B
Next A
Get_Arr_Data_FromTxt = Arr_Data
Erase Arr_Data
End Function
Private Function ExistsFile_UseFso(strPath As String) As Boolean
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
ExistsFile_UseFso = fso.FileExists(strPath)
Set fso = Nothing
End Function
Private Function Transpose_Arr(Arr_Data)
Dim A%, B%
ReDim Arr_output_Temp(LBound(Arr_Data, 2) To UBound(Arr_Data, 2), LBound(Arr_Data, 1) To UBound(Arr_Data, 1))
For A = LBound(Arr_output_Temp, 1) To UBound(Arr_output_Temp, 1)
    For B = LBound(Arr_output_Temp, 2) To UBound(Arr_output_Temp, 2)
        Arr_output_Temp(A, B) = Arr_Data(B, A)
    Next B
Next A
Transpose_Arr = Arr_output_Temp
End Function
Private Sub Get_Arr_StartColumnDateWK()
Dim A%, B&
ReDim Arr_StartColumnDateWK(1 To 3)
For A = UBound(Arr_pitdata, 2) To 9 Step -1
    For B = 2 To UBound(Arr_pitdata)
        If Arr_pitdata(B, 8) = "Inventory On-Hand" Or Arr_pitdata(B, 8) = "Factory On-Hand" Then
            If Not Arr_pitdata(B, A) = Empty Then
                Arr_StartColumnDateWK(1) = A
                Arr_StartColumnDateWK(2) = CDate(Arr_pitdata(1, A))
                Arr_StartColumnDateWK(3) = Get_StartWKfromDate(Arr_StartColumnDateWK(2))
                Exit Sub
            End If
        End If
    Next B
Next A
If Arr_StartColumnDateWK(1) = Empty Then
    Arr_StartColumnDateWK(1) = 10
    Arr_StartColumnDateWK(2) = CDate(Arr_pitdata(1, 10))
    Arr_StartColumnDateWK(3) = Get_StartWKfromDate(Arr_StartColumnDateWK(2))
End If
End Sub
