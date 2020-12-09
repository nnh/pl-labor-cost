Attribute VB_Name = "pl_labor_cost"
Option Explicit
Type pivottable_info
    cst_pivottable_name As String
    dept_idx As Integer
    resource_idx As Integer
    total_spending_idx As Integer
    input_ws As Worksheet
    output_ws As Worksheet
    header_list As Variant
    range_area As String
End Type
Const cst_summary_by_dept As String = "所属毎"
Const cst_summary_by_dept_and_resource As String = "所属・財源毎"
Const cst_fulltime As String = "常勤"
Const cst_parttime As String = "非常勤"
Const cst_full_part As String = "常勤・非常勤"
Const cst_deptlist_name As String = "部署メンバー一覧.xlsx"
Const cst_header_seq As String = "通番"
Const cst_header_id As String = "職員番号"
Const cst_header_name As String = "氏名"
Const cst_header_dept As String = "所属"
Const cst_header_total_spending As String = "総支出額"
Const cst_header_financial_resource As String = "財源"
Const cst_test_f_name As String = "test_f"
Private Function arrayHeaderList() As Variant
Dim cst_header_list As Variant
    cst_header_list = Array("年月", cst_header_seq, cst_header_id, cst_header_name, cst_header_total_spending, cst_header_dept, cst_header_financial_resource)
    arrayHeaderList = cst_header_list
End Function

Public Sub main()
On Error GoTo Finl_L
Const dept_wb_target_ws_count = 2
Dim parentPath As String
Dim inputPath As String
Dim extPath As String
Dim outputPath As String
Dim fileList() As String
Dim temp As String
Dim cnt As Integer
Dim output_wb As Workbook
Dim save_addsheet_cnt As Long
Dim i As Integer
Dim output_last_row As Long
Dim input_str As String
Dim input_year As String
Dim deptlist_wb As Workbook
Dim dept_full_ws_name As String
Dim dept_part_ws_name As String
Dim output_wb_ws_name_list As Variant
Dim add_ws_count As Integer
Dim temp_ws As Worksheet
Dim dept_ws_exist_check As Integer
    Application.ScreenUpdating = False
    save_addsheet_cnt = Application.SheetsInNewWorkbook
    parentPath = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "¥") - 1)
    inputPath = parentPath & "¥input¥rawdata"
    extPath = parentPath & "¥input¥ext"
    outputPath = parentPath & "¥output"
    ' Target all files with "xlsx" extension
    temp = Dir(inputPath & "¥*.xlsx")
    cnt = -1
    Do While temp <> ""
        cnt = cnt + 1
        ReDim Preserve fileList(cnt)
        fileList(cnt) = temp
        temp = Dir()
    Loop
    ' If the file does not exist, exit
    If cnt = -1 Then
        MsgBox prompt:="入力ファイルが存在しなかったため処理を終了します"
        GoTo Finl_L
    End If
    On Error Resume Next
    Workbooks.Open Filename:=extPath & "¥" & cst_deptlist_name
    Set deptlist_wb = Workbooks(cst_deptlist_name)
    On Error GoTo 0
    If deptlist_wb Is Nothing Then
        MsgBox prompt:=cst_deptlist_name & "が存在しなかったため処理を終了します"
        GoTo Finl_L
    End If
    ' Enter password
    input_str = Application.InputBox(prompt:="ファイルのパスワードを入力してください", Type:=2)
    If input_str = "False" Then
        MsgBox prompt:="パスワードが入力されなかったため処理を終了します"
        GoTo Finl_L
    End If
    ' Enter years
    input_year = Application.InputBox(prompt:="処理年度を入力してください", Type:=2)
    If input_year = "False" Then
        MsgBox prompt:="処理年度が入力されなかったため処理を終了します"
        GoTo Finl_L
    End If
    dept_full_ws_name = input_year & cst_fulltime
    dept_part_ws_name = input_year & cst_parttime
    ' Sheet existence check
    dept_ws_exist_check = 0
    For Each temp_ws In deptlist_wb.Worksheets
        If (temp_ws.Name = dept_full_ws_name) Or (temp_ws.Name = dept_part_ws_name) Then
            dept_ws_exist_check = dept_ws_exist_check + 1
        End If
        If dept_ws_exist_check >= dept_wb_target_ws_count Then
            Exit For
        End If
    Next temp_ws
    If dept_ws_exist_check < dept_wb_target_ws_count Then
        MsgBox prompt:="該当年度の部署メンバー一覧シートが存在しなかったため処理を終了します"
        GoTo Finl_L
    End If
    ' Create a workbook for output
    output_wb_ws_name_list = Array(cst_summary_by_dept, cst_summary_by_dept_and_resource, cst_fulltime, cst_parttime, cst_full_part, dept_full_ws_name, dept_part_ws_name)
    add_ws_count = UBound(output_wb_ws_name_list) + 1
    Application.SheetsInNewWorkbook = add_ws_count
    Set output_wb = Workbooks.Add
    For i = LBound(output_wb_ws_name_list) To UBound(output_wb_ws_name_list)
        output_wb.Worksheets(i + 1).Name = output_wb_ws_name_list(i)
    Next i
    If copyFromInputToOutput(inputPath, fileList, output_wb, input_str) Then
        GoTo Err_L
    End If
    Call editOutputWorksheet(output_wb.Worksheets(cst_fulltime))
    Call editOutputWorksheet(output_wb.Worksheets(cst_parttime))
    ' Summary of full-time and part-time
    output_last_row = copyAllCells(output_wb.Worksheets(cst_fulltime), output_wb.Worksheets(cst_full_part), 1, 1, 1, 1)
    output_last_row = copyAllCells(output_wb.Worksheets(cst_parttime), output_wb.Worksheets(cst_full_part), output_last_row + 1, 1, 2, 1)
    Application.DisplayAlerts = False
    output_wb.Worksheets(cst_fulltime).Delete
    output_wb.Worksheets(cst_parttime).Delete
    Application.DisplayAlerts = True
    Call sortCellsBySeqAndYear(output_wb.Worksheets(cst_full_part), xlYes)
    ' Link department
    Call linkDepartmentByName(output_wb.Worksheets(cst_full_part), deptlist_wb, dept_full_ws_name, dept_part_ws_name)
    ' Total Expenditure for each staff number
    Call createPivottableByDept(output_wb, cst_full_part, cst_summary_by_dept)
    Call createPivottableByDeptAndResource(output_wb, cst_full_part, cst_summary_by_dept_and_resource)
    ' Save the output workbook
    With output_wb
        With Worksheets(1)
            .Activate
            .Cells(1, 1).Select
        End With
        .SaveAs outputPath & "¥" & Format(Now(), "yyyymmdd_hhmmss") & ".xlsx"
        .Close
    End With
    GoTo Finl_L
Err_L:
    output_wb.Close savechanges:=False
Finl_L:
    deptlist_wb.Close savechanges:=False
    Application.SheetsInNewWorkbook = save_addsheet_cnt
    Application.ScreenUpdating = True
    MsgBox prompt:="処理が終了しました"
End Sub

Private Function copyDeptWs(copyFrom_wb As Workbook, copyTo_ws As Worksheet) As Worksheet
Dim temp As Long
Dim dept_ws_name As String
    dept_ws_name = copyTo_ws.Name
    temp = copyAllCells(copyFrom_wb.Worksheets(dept_ws_name), copyTo_ws, 1, 1, 1, 1)
    Set copyDeptWs = copyTo_ws
End Function

Private Sub linkDepartmentByName(target_ws As Worksheet, deptlist_wb As Workbook, dept_full_ws_name As String, dept_part_ws_name As String)
Const dept_vlookup_idx As Integer = 3
Const financial_resource_vlookup_idx As Integer = 8
Const vlookupRange As String = "B:I"
Dim target_header_list() As Variant
Dim last_row As Long
Dim i As Long
Dim input_str As String
Dim input_lng As Long
Dim target_id_idx As Integer
Dim target_name_idx As Integer
Dim target_output_idx As Integer
Dim target_output_financial_resource_idx As Integer
Dim dept_info As Variant
Dim financial_resource_info As Variant
Dim dept_full_ws As Worksheet
Dim dept_part_ws As Worksheet
Dim output_wb As Workbook
    Set output_wb = target_ws.Parent
    Set dept_full_ws = copyDeptWs(deptlist_wb, output_wb.Worksheets(dept_full_ws_name))
    Set dept_part_ws = copyDeptWs(deptlist_wb, output_wb.Worksheets(dept_part_ws_name))
    target_header_list = arrayHeaderList()
    target_id_idx = getArrayIdx(target_header_list, cst_header_id)
    target_name_idx = getArrayIdx(target_header_list, cst_header_name)
    target_output_idx = getArrayIdx(target_header_list, cst_header_dept)
    target_output_financial_resource_idx = getArrayIdx(target_header_list, cst_header_financial_resource)
    last_row = target_ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    For i = 2 To last_row
        input_str = target_ws.Cells(i, target_id_idx + 1).Value
        If (input_str = "") Or (Not IsNumeric(input_str)) Then
            Exit For
        End If
        input_lng = CLng(input_str)
        dept_info = exec_vlookup(input_lng, dept_full_ws.Range(vlookupRange), dept_vlookup_idx)
        financial_resource_info = exec_vlookup(input_lng, dept_full_ws.Range(vlookupRange), financial_resource_vlookup_idx)
        If IsEmpty(dept_info) Then
            dept_info = exec_vlookup(input_lng, dept_part_ws.Range(vlookupRange), dept_vlookup_idx)
            financial_resource_info = exec_vlookup(input_lng, dept_part_ws.Range(vlookupRange), financial_resource_vlookup_idx)
        End If
        If IsEmpty(dept_info) Then
            dept_info = "！！！エラー！！！"
            financial_resource_info = ""
        Else
        End If
        target_ws.Cells(i, target_output_idx + 1) = dept_info
        target_ws.Cells(i, target_output_financial_resource_idx + 1) = financial_resource_info
        
        If ThisWorkbook.Worksheets(1).CheckBoxes(cst_test_f_name) = xlOff Then
            target_ws.Cells(i, target_id_idx + 1) = ""
        End If
        target_ws.Cells(i, target_name_idx + 1) = ""
    Next i
    dept_full_ws.Visible = xlSheetHidden
    dept_part_ws.Visible = xlSheetHidden
    Set dept_part_ws = Nothing
    Set dept_full_ws = Nothing
End Sub

Private Function exec_vlookup(input_lng As Long, vloookup_range As Range, target_idx As Integer) As Variant
Dim temp As Variant
    On Error Resume Next
    temp = Application.WorksheetFunction.VLookup(input_lng, vloookup_range, target_idx, False)
    On Error GoTo 0
    exec_vlookup = temp
End Function

Private Function copyFromInputToOutput(input_path As String, file_list() As String, output_wb As Workbook, input_password As String) As Boolean
On Error GoTo Finl_L
Dim input_wb As Workbook
Dim ws As Sheets
Dim temp_ws As Worksheet
Dim yymm As String
Dim output_fulltime_last_row As Long
Dim output_parttime_last_row As Long
Dim i As Long
Dim file_Path As String
Dim error_str As String
Dim error_f As Boolean
    error_f = False
    ThisWorkbook.Worksheets(1).Cells.Clear
    For i = LBound(file_list) To UBound(file_list)
        On Error Resume Next
        Set input_wb = Workbooks.Open(Filename:=input_path & "¥" & file_list(i), ReadOnly:=True, password:=input_password)
        If Err.Number = 0 Then
            error_str = error_str & file_list(i) & "をオープンしました" & vbCrLf
        Else
            error_str = error_str & file_list(i) & "のオープンに失敗しました" & vbCrLf
            MsgBox "ファイルのオープンに失敗したため処理を終了します"
            error_f = True
            GoTo Finl_L
        End If
        On Error GoTo 0
        yymm = "'" & Left(input_wb.Name, 4)
        If InStr(input_wb.Name, "賞与") > 0 Then
            yymm = yymm & "_賞与"
        End If
        Set ws = input_wb.Worksheets
        For Each temp_ws In ws
            If Left(temp_ws.Name, Len(cst_fulltime)) = cst_fulltime Then
                output_fulltime_last_row = outputValues(temp_ws, output_wb.Worksheets(cst_fulltime), output_fulltime_last_row, yymm)
            End If
            If Left(temp_ws.Name, Len(cst_parttime)) = cst_parttime Then
                output_parttime_last_row = outputValues(temp_ws, output_wb.Worksheets(cst_parttime), output_parttime_last_row, yymm)
            End If
        Next temp_ws
        input_wb.Close savechanges:=False
    Next i
Finl_L:
    ThisWorkbook.Worksheets(1).Cells(1, 1) = Left(error_str, Len(error_str) - 1)
    Set ws = Nothing
    Set input_wb = Nothing
    copyFromInputToOutput = error_f
End Function

Private Function outputValues(input_ws As Worksheet, output_ws As Worksheet, save_last_row As Long, yymm As String) As Long
Const cst_input_start_row As Integer = 4
Const cst_input_start_col As Integer = 1
Const cst_output_start_col As Integer = 2
Dim output_start_row As Long
Dim output_last_row As Long
Dim i As Long
    output_start_row = save_last_row + 1
    output_last_row = copyAllCells(input_ws, output_ws, output_start_row, cst_output_start_col, cst_input_start_row, cst_input_start_col)
    ' Column A : year and month
    For i = output_start_row To output_last_row
        output_ws.Cells(i, 1) = yymm
    Next i
    outputValues = output_last_row

End Function

Private Sub editOutputWorksheet(output_ws As Worksheet)
Const cst_fulltime_total_spending_col As String = "AO"
Const cst_parttime_total_spending_col As String = "Z"
Dim total_spending_col As String
Dim blank_address As String
Dim target_rows As String
Dim last_col As String
Dim one_before_total_spending_col As String
Dim one_after_total_spending_col As String
Dim temp_col As String
Dim header_list() As Variant
Dim i As Integer
Dim temp_row As Long
    header_list = arrayHeaderList()
    If output_ws.Name = cst_fulltime Then
        total_spending_col = cst_fulltime_total_spending_col
    Else
        total_spending_col = cst_parttime_total_spending_col
    End If
    output_ws.Activate
    ' If column C is blank, delete the row
    output_ws.Cells.Sort key1:=Range("C1"), Header:=xlNo
    blank_address = output_ws.Range("C:C").SpecialCells(xlCellTypeBlanks).Address
    target_rows = Replace(blank_address, "$C$", "")
    output_ws.Rows(target_rows).Delete
    ' Sort by year and month, seq
    Call sortCellsBySeqAndYear(output_ws, xlNo)
    ' Remove columns, leaving columns A-D and total expenditures
    one_before_total_spending_col = convertToLetter(output_ws.Cells.Range(total_spending_col & "1").Column - 1)
    one_after_total_spending_col = convertToLetter(output_ws.Cells.Range(total_spending_col & "1").Column + 1)
    ' Get the last column
    last_col = convertToLetter(output_ws.Cells.SpecialCells(xlCellTypeLastCell).Column)
    output_ws.Columns(one_after_total_spending_col & ":" & last_col).Delete
    output_ws.Columns("E:" & one_before_total_spending_col).Delete
    ' Add headline
    output_ws.Rows(1).Insert
    For i = LBound(header_list) To UBound(header_list)
        output_ws.Cells(1, i + 1) = header_list(i)
    Next i
    ' Edit seq
Dim target_header_list() As Variant
Dim target_seq_idx As Integer
Dim temp_str As String
    target_header_list = arrayHeaderList()
    target_seq_idx = getArrayIdx(target_header_list, cst_header_seq)
    temp_row = 2
    Do
        temp_str = Trim(output_ws.Cells(temp_row, target_seq_idx + 1).Value)
        If temp_str = "" Or Not IsNumeric(temp_str) Then
            Exit Do
        End If
        output_ws.Cells(temp_row, target_seq_idx + 1).Value = output_ws.Name & "_" & Format(temp_str, "000")
        temp_row = temp_row + 1
    Loop
    
End Sub

' https://docs.microsoft.com/ja-jp/office/troubleshoot/excel/convert-excel-column-numbers
Private Function convertToLetter(iCol As Long) As String
Dim a As Long
Dim b As Long
    a = iCol
    convertToLetter = ""
    Do While iCol > 0
        a = Int((iCol - 1) / 26)
        b = (iCol - 1) Mod 26
        convertToLetter = Chr(b + 65) & convertToLetter
        iCol = a
    Loop
End Function

Private Function copyAllCells(input_ws As Worksheet, output_ws As Worksheet, output_start_row As Long, output_start_col As Long, input_start_row As Long, input_start_col As Long)
Dim input_last_cell As Range
Dim output_last_row As Long
    If input_ws.FilterMode Then
        input_ws.ShowAllData
    End If
    Set input_last_cell = input_ws.Cells.SpecialCells(xlCellTypeLastCell)
    output_last_row = output_start_row + input_last_cell.Row - input_start_row
    input_ws.Range(input_ws.Cells(input_start_row, input_start_col), input_last_cell).Copy
    output_ws.Cells(output_start_row, output_start_col).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    output_ws.Activate
    output_ws.Range("A1").Select
    copyAllCells = output_last_row
End Function

Private Sub sortCellsBySeqAndYear(target_ws As Worksheet, header_f As Variant)
    target_ws.Activate
    target_ws.Cells.Sort key1:=Range("A1"), order1:=xlAscending, dataoption1:=xlSortTextAsNumbers, _
                         key2:=Range("B1"), order2:=xlAscending, dataoption2:=xlSortTextAsNumbers, _
                         Header:=header_f, Orientation:=xlSortColumns, SortMethod:=xlStroke
End Sub

Private Function getArrayIdx(array_list As Variant, target_str As String) As Integer
Dim temp As Variant
Dim res As Integer
    On Error Resume Next
    temp = Application.WorksheetFunction.Match(target_str, array_list, 0)
    On Error GoTo 0
    If IsEmpty(temp) Then
        res = -1
    Else
        res = temp - 1
    End If
    getArrayIdx = res
End Function

Private Sub createPivottableByDept(output_wb As Workbook, input_ws_name As String, output_ws_name As String)
Dim pv_info As pivottable_info
    pv_info = setPivottableInfo(output_wb, input_ws_name, output_ws_name, "pv2")
    output_wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pv_info.input_ws.Range(pv_info.range_area)).createPivottable _
                                 TableDestination:=pv_info.output_ws.Range("A1"), TableName:=pv_info.cst_pivottable_name
    With pv_info.output_ws.PivotTables(pv_info.cst_pivottable_name)
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
        .AddDataField pv_info.output_ws.PivotTables(pv_info.cst_pivottable_name).PivotFields(pv_info.header_list(pv_info.dept_idx))
        .PivotFields(pv_info.header_list(pv_info.dept_idx)).Orientation = xlRowField
        .AddDataField pv_info.output_ws.PivotTables(pv_info.cst_pivottable_name).PivotFields(pv_info.header_list(pv_info.total_spending_idx)), "合計 / " & pv_info.header_list(pv_info.total_spending_idx), xlSum
'        .PivotFields("個数 / " & cst_header_dept).Orientation = xlHidden
    End With
End Sub
Private Sub createPivottableByDeptAndResource(output_wb As Workbook, input_ws_name As String, output_ws_name As String)
Dim pv_info As pivottable_info
    pv_info = setPivottableInfo(output_wb, input_ws_name, output_ws_name, "pv3")
    output_wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pv_info.input_ws.Range(pv_info.range_area)).createPivottable _
                                 TableDestination:=pv_info.output_ws.Range("A1"), TableName:=pv_info.cst_pivottable_name
    With pv_info.output_ws.PivotTables(pv_info.cst_pivottable_name)
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
        .AddDataField pv_info.output_ws.PivotTables(pv_info.cst_pivottable_name).PivotFields(pv_info.header_list(pv_info.dept_idx))
        .PivotFields(pv_info.header_list(pv_info.dept_idx)).Orientation = xlRowField
        .AddDataField pv_info.output_ws.PivotTables(pv_info.cst_pivottable_name).PivotFields(pv_info.header_list(pv_info.resource_idx))
        .PivotFields(pv_info.header_list(pv_info.resource_idx)).Orientation = xlRowField
        .AddDataField pv_info.output_ws.PivotTables(pv_info.cst_pivottable_name).PivotFields(pv_info.header_list(pv_info.total_spending_idx)), "合計 / " & pv_info.header_list(pv_info.total_spending_idx), xlSum
'        .PivotFields("個数 / " & cst_header_dept).Orientation = xlHidden
'        .PivotFields("個数 / " & cst_header_financial_resource).Orientation = xlHidden
        .PivotFields(cst_header_dept).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
End Sub
Private Function setPivottableInfo(output_wb As Workbook, input_ws_name As String, output_ws_name As String, pv_name As String) As pivottable_info
Dim pv_info As pivottable_info
    output_wb.Worksheets(output_ws_name).Activate
    With pv_info
        .cst_pivottable_name = pv_name
        .header_list = arrayHeaderList()
        .dept_idx = getArrayIdx(.header_list, cst_header_dept)
        .resource_idx = getArrayIdx(.header_list, cst_header_financial_resource)
        .total_spending_idx = getArrayIdx(.header_list, cst_header_total_spending)
        Set .input_ws = output_wb.Worksheets(input_ws_name)
        Set .output_ws = output_wb.Worksheets(output_ws_name)
        .range_area = "A:G"
    End With
    setPivottableInfo = pv_info
End Function
