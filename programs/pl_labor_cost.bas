Attribute VB_Name = "pl_labor_cost"
Option Explicit
Const cst_summary As String = "summary"
Const cst_fulltime As String = "常勤"
Const cst_parttime As String = "非常勤"
Const cst_full_part As String = "常勤・非常勤"
Private Function arrayHeaderList() As Variant
Dim cst_header_list As Variant
    cst_header_list = Array("年月", "通番", "職員番号", "氏名", "総支出額")
    arrayHeaderList = cst_header_list
End Function

Public Sub main()
On Error GoTo Finl_L
Dim parentPath As String
Dim inputPath As String
Dim outputPath As String
Dim fileList() As String
Dim temp As String
Dim cnt As Integer
Dim output_wb As Workbook
Dim save_addsheet_cnt As Long
Dim i As Integer
Dim output_last_row As Long
Dim input_str As String
    Application.ScreenUpdating = False
    save_addsheet_cnt = Application.SheetsInNewWorkbook
    parentPath = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\") - 1)
    inputPath = parentPath & "\input"
    outputPath = parentPath & "\output"
    ' Target all files with "xlsx" extension
    temp = Dir(inputPath & "\*.xlsx")
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
    ' Enter password
    input_str = Application.InputBox(prompt:="ファイルのパスワードを入力してください", Type:=2)
    If input_str = "False" Then
        MsgBox prompt:="パスワードが入力されなかったため処理を終了します"
        GoTo Finl_L
    End If
    ' Create a workbook for output
    Application.SheetsInNewWorkbook = 4
    Set output_wb = Workbooks.Add
    output_wb.Worksheets(1).Name = cst_summary
    output_wb.Worksheets(2).Name = cst_fulltime
    output_wb.Worksheets(3).Name = cst_parttime
    output_wb.Worksheets(4).Name = cst_full_part
    Call copyFromInputToOutput(inputPath, fileList, output_wb, input_str)
    Call editOutputWorksheet(output_wb.Worksheets(cst_fulltime))
    Call editOutputWorksheet(output_wb.Worksheets(cst_parttime))
    ' Summary of full-time and part-time
    output_last_row = copyAllCells(output_wb.Worksheets(cst_fulltime), output_wb.Worksheets(cst_full_part), 1, 1, 1, 1)
    output_last_row = copyAllCells(output_wb.Worksheets(cst_parttime), output_wb.Worksheets(cst_full_part), output_last_row + 1, 1, 2, 1)
    Call sortCellsBySeqAndYear(output_wb.Worksheets(cst_full_part), xlYes)
    ' Total Expenditure for each staff number
    output_wb.Worksheets(cst_summary).Activate
    Call createPivottable(output_wb, cst_full_part, cst_summary)
    output_wb.SaveAs outputPath & "\" & Format(Now(), "yyyymmddhhmmss") & "test.xlsx"
    output_wb.Close
Finl_L:
    Application.SheetsInNewWorkbook = save_addsheet_cnt
    Application.ScreenUpdating = True
    MsgBox prompt:="処理が終了しました"
End Sub

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
        Set input_wb = Workbooks.Open(Filename:=input_path & "\" & file_list(i), ReadOnly:=True, password:=input_password)
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
        input_wb.Close savechanges:=False
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

Private Sub createPivottable(output_wb As Workbook, input_ws_name As String, output_ws_name As String)
Const cst_pivottable_name As String = "pt1"
Const cst_id_idx As Integer = 2
Const cst_name_idx As Integer = 3
Const cst_total_spending_idx As Integer = 4
Dim input_ws As Worksheet
Dim output_ws As Worksheet
Dim header_list As Variant
    header_list = arrayHeaderList()
    Set input_ws = output_wb.Worksheets(input_ws_name)
    Set output_ws = output_wb.Worksheets(output_ws_name)
    output_wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=input_ws.Range("A:E")).createPivottable _
                                 TableDestination:=output_ws.Range("A1"), TableName:=cst_pivottable_name
    With output_ws.PivotTables(cst_pivottable_name)
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
        .AddDataField output_ws.PivotTables(cst_pivottable_name).PivotFields(header_list(cst_id_idx))
        .PivotFields(header_list(cst_id_idx)).Orientation = xlRowField
        .AddDataField output_ws.PivotTables(cst_pivottable_name).PivotFields(header_list(cst_name_idx))
        .PivotFields(header_list(cst_name_idx)).Orientation = xlRowField
        .AddDataField output_ws.PivotTables(cst_pivottable_name).PivotFields(header_list(cst_total_spending_idx)), "合計 / " & header_list(cst_total_spending_idx), xlSum
        .PivotFields(header_list(cst_id_idx)).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields("合計 / " & header_list(cst_id_idx)).Orientation = xlHidden
        .PivotFields("個数 / " & header_list(cst_name_idx)).Orientation = xlHidden
    End With
    Set output_ws = Nothing
    Set input_ws = Nothing
End Sub
