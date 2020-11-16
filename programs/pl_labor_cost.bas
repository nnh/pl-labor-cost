Attribute VB_Name = "pl_labor_cost"
Option Explicit
Const cst_summary_by_dept As String = "所属毎"
Const cst_summary_by_dept_and_resource As String = "所属・財源毎"
Const cst_fulltime As String = "常勤"
Const cst_parttime As String = "非常勤"
Const cst_full_part As String = "常勤・非常勤"
Const cst_deptlist_name As String = "部署メンバー一覧.xlsx"
Const cst_header_id As String = "職員番号"
Const cst_header_name As String = "氏名"
Const cst_header_dept As String = "所属"
Const cst_header_total_spending As String = "総支出額"
Const cst_header_financial_resource As String = "財源"
Private Function arrayHeaderList() As Variant
Dim cst_header_list As Variant
    cst_header_list = Array("年月", "通番", "職員番号", cst_header_name, cst_header_total_spending, cst_header_dept, cst_header_financial_resource)
    arrayHeaderList = cst_header_list
End Function

Public Sub main()
On Error GoTo Finl_L
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
Dim dept_full_ws As Worksheet
Dim dept_part_ws As Worksheet
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
    ' Create a workbook for output
    Application.SheetsInNewWorkbook = 5
    Set output_wb = Workbooks.Add
    output_wb.Worksheets(1).Name = cst_summary_by_dept
    output_wb.Worksheets(2).Name = cst_summary_by_dept_and_resource
    output_wb.Worksheets(3).Name = cst_fulltime
    output_wb.Worksheets(4).Name = cst_parttime
    output_wb.Worksheets(5).Name = cst_full_part
    Call copyFromInputToOutput(inputPath, fileList, output_wb, input_str)
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
    Set dept_full_ws = copyDeptWs(deptlist_wb, output_wb, input_year & cst_fulltime)
    Set dept_part_ws = copyDeptWs(deptlist_wb, output_wb, input_year & cst_parttime)
    Call linkDepartmentByName(output_wb.Worksheets(cst_full_part), dept_full_ws, dept_part_ws)
    ' Total Expenditure for each staff number
    output_wb.Worksheets(cst_summary_by_dept).Activate
    Call createPivottableByDept(output_wb, cst_full_part, cst_summary_by_dept)
    output_wb.Worksheets(cst_summary_by_dept_and_resource).Activate
    Call createPivottableByDeptAndResource(output_wb, cst_full_part, cst_summary_by_dept_and_resource)
    dept_full_ws.Visible = xlSheetHidden
    dept_part_ws.Visible = xlSheetHidden
    With output_wb.Worksheets(1)
        .Activate
        .Cells(1, 1).Select
    End With
    output_wb.SaveAs outputPath & "¥" & Format(Now(), "yyyymmdd_hhmmss") & ".xlsx"
    output_wb.Close
Finl_L:
    deptlist_wb.Close saveChanges:=False
    Application.SheetsInNewWorkbook = save_addsheet_cnt
    Application.ScreenUpdating = True
    MsgBox prompt:="処理が終了しました"
End Sub

Private Function copyDeptWs(copyFrom_wb As Workbook, copyTo_wb As Workbook, dept_ws_name As String) As Worksheet
Dim dept_ws As Worksheet
    copyFrom_wb.Worksheets(dept_ws_name).Copy after:=copyTo_wb.Worksheets(copyTo_wb.Worksheets.Count)
    Set dept_ws = copyTo_wb.Worksheets(dept_ws_name)
    Call removeCellsBlank(dept_ws)
    Set copyDeptWs = dept_ws
End Function

Private Sub removeCellsBlank(target_ws As Worksheet)
Const cst_target_col As Integer = 2
Dim last_row As Long
Dim i As Long
Dim temp_str As String
    last_row = target_ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    For i = 3 To last_row
        If target_ws.Cells(i, cst_target_col).Value = "" Then
            Exit For
        End If
        target_ws.Cells(i, cst_target_col).Value = removeBlank(target_ws.Cells(i, cst_target_col).Value)
    Next i
End Sub

Private Function removeBlank(target_value As String) As String
Dim temp_str As String
    temp_str = Replace(target_value, "　", "")
    temp_str = Replace(temp_str, " ", "")
    removeBlank = temp_str
End Function

Private Sub linkDepartmentByName(target_ws As Worksheet, fulltime_dept_ws As Worksheet, parttime_dept_ws As Worksheet)
Const dept_vlookup_idx As Integer = 2
Const financial_resource_vlookup_idx As Integer = 6
Const vlookupRange As String = "B:H"
Dim target_header_list() As Variant
Dim last_row As Long
Dim i As Long
Dim input_str As String
Dim target_name_idx As Integer
Dim target_output_idx As Integer
Dim target_output_financial_resource_idx As Integer
Dim dept_info As Variant
Dim financial_resource_info As Variant
    target_header_list = arrayHeaderList()
    target_name_idx = getArrayIdx(target_header_list, cst_header_name)
    target_output_idx = getArrayIdx(target_header_list, cst_header_dept)
    target_output_financial_resource_idx = getArrayIdx(target_header_list, cst_header_financial_resource)
    last_row = target_ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    For i = 2 To last_row
        input_str = target_ws.Cells(i, target_name_idx + 1).Value
        If input_str = "" Then
            Exit For
        End If
        dept_info = exec_vlookup(input_str, fulltime_dept_ws.Range(vlookupRange), dept_vlookup_idx)
        financial_resource_info = exec_vlookup(input_str, fulltime_dept_ws.Range(vlookupRange), financial_resource_vlookup_idx)
        If IsEmpty(dept_info) Then
            dept_info = exec_vlookup(input_str, parttime_dept_ws.Range(vlookupRange), dept_vlookup_idx)
            financial_resource_info = exec_vlookup(input_str, parttime_dept_ws.Range(vlookupRange), financial_resource_vlookup_idx)
        End If
        If IsEmpty(dept_info) Then
            dept_info = "！！！エラー！！！"
            financial_resource_info = ""
        End If
        target_ws.Cells(i, target_output_idx + 1) = dept_info
        target_ws.Cells(i, target_output_financial_resource_idx + 1) = financial_resource_info
        target_ws.Cells(i, target_name_idx + 1) = ""
    Next i
End Sub

Private Function exec_vlookup(input_str As String, vloookup_range As Range, target_idx As Integer) As Variant
Dim temp As Variant
Dim temp_str As String
    temp_str = removeBlank(input_str)
    On Error Resume Next
    temp = Application.WorksheetFunction.VLookup(temp_str, vloookup_range, target_idx, False)
    On Error GoTo 0
    exec_vlookup = temp
End Function

Private Sub copyFromInputToOutput(input_path As String, file_list() As String, output_wb As Workbook, input_password As String)
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
    ThisWorkbook.Worksheets(1).Cells.Clear
    For i = LBound(file_list) To UBound(file_list)
        On Error Resume Next
        Set input_wb = Workbooks.Open(Filename:=input_path & "¥" & file_list(i), ReadOnly:=True, password:=input_password)
        If Err.Number = 0 Then
            error_str = error_str & file_list(i) & "をオープンしました" & vbCrLf
        Else
            error_str = error_str & file_list(i) & "のオープンに失敗しました" & vbCrLf
            Err.Clear
            GoTo nextfile
        End If
        On Error GoTo 0
        yymm = "'" & Left(input_wb.Name, 4)
        Set ws = input_wb.Worksheets
        For Each temp_ws In ws
            If Left(temp_ws.Name, Len(cst_fulltime)) = cst_fulltime Then
                output_fulltime_last_row = outputValues(temp_ws, output_wb.Worksheets(cst_fulltime), output_fulltime_last_row, yymm)
            End If
            If Left(temp_ws.Name, Len(cst_parttime)) = cst_parttime Then
                output_parttime_last_row = outputValues(temp_ws, output_wb.Worksheets(cst_parttime), output_parttime_last_row, yymm)
            End If
        Next temp_ws
        input_wb.Close saveChanges:=False
nextfile:
    Next i
Finl_L:
    ThisWorkbook.Worksheets(1).Cells(1, 1) = Left(error_str, Len(error_str) - 1)
    Set ws = Nothing
    Set input_wb = Nothing
End Sub

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
Const cst_pivottable_name As String = "pt2"
Dim dept_idx As Integer
Dim total_spending_idx As Integer
Dim input_ws As Worksheet
Dim output_ws As Worksheet
Dim header_list As Variant
    header_list = arrayHeaderList()
    dept_idx = getArrayIdx(header_list, cst_header_dept)
    total_spending_idx = getArrayIdx(header_list, cst_header_total_spending)
    Set input_ws = output_wb.Worksheets(input_ws_name)
    Set output_ws = output_wb.Worksheets(output_ws_name)
    output_wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=input_ws.Range("A:G")).createPivottable _
                                 TableDestination:=output_ws.Range("A1"), TableName:=cst_pivottable_name
    With output_ws.PivotTables(cst_pivottable_name)
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
        .AddDataField output_ws.PivotTables(cst_pivottable_name).PivotFields(header_list(dept_idx))
        .PivotFields(header_list(dept_idx)).Orientation = xlRowField
        .AddDataField output_ws.PivotTables(cst_pivottable_name).PivotFields(header_list(total_spending_idx)), "合計 / " & header_list(total_spending_idx), xlSum
        .PivotFields("個数 / " & cst_header_dept).Orientation = xlHidden
    End With
    Set output_ws = Nothing
    Set input_ws = Nothing
End Sub
Private Sub createPivottableByDeptAndResource(output_wb As Workbook, input_ws_name As String, output_ws_name As String)
Const cst_pivottable_name As String = "pt3"
Dim dept_idx As Integer
Dim resource_idx As Integer
Dim total_spending_idx As Integer
Dim input_ws As Worksheet
Dim output_ws As Worksheet
Dim header_list As Variant
    header_list = arrayHeaderList()
    dept_idx = getArrayIdx(header_list, cst_header_dept)
    resource_idx = getArrayIdx(header_list, cst_header_financial_resource)
    total_spending_idx = getArrayIdx(header_list, cst_header_total_spending)
    Set input_ws = output_wb.Worksheets(input_ws_name)
    Set output_ws = output_wb.Worksheets(output_ws_name)
    output_wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=input_ws.Range("A:G")).createPivottable _
                                 TableDestination:=output_ws.Range("A1"), TableName:=cst_pivottable_name
    With output_ws.PivotTables(cst_pivottable_name)
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
        .AddDataField output_ws.PivotTables(cst_pivottable_name).PivotFields(header_list(dept_idx))
        .PivotFields(header_list(dept_idx)).Orientation = xlRowField
        .AddDataField output_ws.PivotTables(cst_pivottable_name).PivotFields(header_list(resource_idx))
        .PivotFields(header_list(resource_idx)).Orientation = xlRowField
        .AddDataField output_ws.PivotTables(cst_pivottable_name).PivotFields(header_list(total_spending_idx)), "合計 / " & header_list(total_spending_idx), xlSum
        .PivotFields("個数 / " & cst_header_dept).Orientation = xlHidden
        .PivotFields("個数 / " & cst_header_financial_resource).Orientation = xlHidden
        .PivotFields(cst_header_dept).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    Set output_ws = Nothing
    Set input_ws = Nothing
End Sub
