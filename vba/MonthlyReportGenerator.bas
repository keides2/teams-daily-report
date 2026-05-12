Attribute VB_Name = "MonthlyReportGenerator"
Option Explicit

Private Const TEMPLATE_SHEET_NAME As String = "実施申請・報告書（2026年4月）"
Private Const FISCAL_START_YEAR As Long = 2026
Private Const FIRST_BLOCK_ROW As Long = 21
Private Const BLOCK_HEIGHT As Long = 8
Private Const LAST_POSSIBLE_DAY As Long = 31

Public Sub GenerateMonthlyReportFiles()
    Dim outputFolder As String
    Dim currentWorkbookPath As String
    Dim targetMonth As Long
    Dim reportYear As Long
    Dim fileName As String
    Dim fullPath As String
    Dim wbCopy As Workbook
    Dim generatedCount As Long
    Dim usedCurrentWorkbook As Boolean
    Dim lastAction As String

    If Not SheetExists(ThisWorkbook, TEMPLATE_SHEET_NAME) Then
        MsgBox "テンプレートシートが見つかりません: " & TEMPLATE_SHEET_NAME, vbExclamation
        Exit Sub
    End If

    outputFolder = GetOutputFolderPath()
    If outputFolder = "" Then
        MsgBox "保存先フォルダーを取得できませんでした。", vbCritical
        Exit Sub
    End If
    If Right$(outputFolder, 1) <> "\" Then outputFolder = outputFolder & "\"

    currentWorkbookPath = ThisWorkbook.FullName

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    For targetMonth = 1 To 12
        reportYear = GetReportYear(targetMonth)
        fileName = "（" & targetMonth & "月_嶋谷圭介）業務内容報告書.xlsm"
        fullPath = outputFolder & fileName

        If PathsEqual(fullPath, currentWorkbookPath) Then
            lastAction = "実行中ブックを " & targetMonth & "月分として整形中"
            PrepareCurrentWorkbook ThisWorkbook, reportYear, targetMonth
            usedCurrentWorkbook = True
            generatedCount = generatedCount + 1
            GoTo ContinueLoop
        End If

        If FileExists(fullPath) Then
            lastAction = "既存ファイル削除中: " & fullPath
            If Not DeleteIfExists(fullPath) Then
                Err.Raise vbObjectError + 1000, , "既存ファイルを削除できませんでした: " & fullPath
            End If
        End If

        lastAction = "テンプレートのコピー作成中: " & fullPath
        ThisWorkbook.SaveCopyAs fullPath

        lastAction = "コピーしたブックを開いて加工中: " & fullPath
        Set wbCopy = Workbooks.Open(fileName:=fullPath, UpdateLinks:=0, ReadOnly:=False)
        PrepareGeneratedWorkbook wbCopy, reportYear, targetMonth
        wbCopy.Close SaveChanges:=True
        Set wbCopy = Nothing

        generatedCount = generatedCount + 1
ContinueLoop:
    Next targetMonth

CleanExit:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If generatedCount = 12 Then
        If usedCurrentWorkbook Then
            MsgBox "12か月分のファイルを準備しました。" & vbCrLf & _
                   "4月分は実行中のブック自身を使用しました。" & vbCrLf & _
                   "保存先: " & outputFolder, vbInformation
        Else
            MsgBox "12か月分のファイルを生成しました。" & vbCrLf & _
                   "保存先: " & outputFolder, vbInformation
        End If
    End If
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not wbCopy Is Nothing Then wbCopy.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "月次ファイル生成中にエラーが発生しました。" & vbCrLf & _
           "対象月: " & targetMonth & "月" & vbCrLf & _
           "処理: " & lastAction & vbCrLf & _
           "内容: " & Err.Number & " / " & Err.Description, vbCritical
End Sub

Private Sub PrepareCurrentWorkbook(ByVal wb As Workbook, ByVal reportYear As Long, ByVal reportMonth As Long)
    Dim ws As Worksheet
    Dim expectedSheetName As String

    expectedSheetName = "実施申請・報告書（" & reportYear & "年" & reportMonth & "月）"

    If Not SheetExists(wb, TEMPLATE_SHEET_NAME) Then
        If SheetExists(wb, expectedSheetName) Then
            Set ws = wb.Worksheets(expectedSheetName)
        Else
            Err.Raise vbObjectError + 1002, , "実行中ブックに対象シートが見つかりません。"
        End If
    Else
        Set ws = wb.Worksheets(TEMPLATE_SHEET_NAME)
        If ws.Name <> expectedSheetName Then
            ws.Name = expectedSheetName
        End If
    End If

    ApplyMonthlyLayoutToSheet ws, reportYear, reportMonth
    wb.Save
End Sub

Private Sub PrepareGeneratedWorkbook(ByVal wb As Workbook, ByVal reportYear As Long, ByVal reportMonth As Long)
    Dim i As Long
    Dim ws As Worksheet
    Dim targetSheetName As String

    targetSheetName = "実施申請・報告書（" & reportYear & "年" & reportMonth & "月）"

    If Not SheetExists(wb, TEMPLATE_SHEET_NAME) Then
        Err.Raise vbObjectError + 1001, , "コピー先ブックにテンプレートシートがありません。"
    End If

    For i = wb.Worksheets.Count To 1 Step -1
        If wb.Worksheets(i).Name <> TEMPLATE_SHEET_NAME Then
            wb.Worksheets(i).Delete
        End If
    Next i

    Set ws = wb.Worksheets(TEMPLATE_SHEET_NAME)
    ws.Name = targetSheetName

    ApplyMonthlyLayoutToSheet ws, reportYear, reportMonth
    wb.Save
End Sub

Private Sub ApplyMonthlyLayoutToSheet(ByVal ws As Worksheet, ByVal reportYear As Long, ByVal reportMonth As Long)
    Dim lastDay As Long
    Dim dayIndex As Long
    Dim rowNo As Long
    Dim d As Date

    lastDay = Day(DateSerial(reportYear, reportMonth + 1, 0))

    For dayIndex = 1 To LAST_POSSIBLE_DAY
        rowNo = FIRST_BLOCK_ROW + (dayIndex - 1) * BLOCK_HEIGHT

        If dayIndex <= lastDay Then
            d = DateSerial(reportYear, reportMonth, dayIndex)

            ClearCellSafe ws.Cells(rowNo, "B")
            ClearCellSafe ws.Cells(rowNo + 1, "B")
            SetCellValueSafe ws.Cells(rowNo + 3, "B"), Format$(d, "mm/dd")
            SetCellValueSafe ws.Cells(rowNo + 4, "B"), "(" & Format$(d, "aaa") & ")"

            If IsWeekend(d) Then
                SetCellValueSafe ws.Cells(rowNo, "C"), "-"
                SetCellValueSafe ws.Cells(rowNo, "E"), "-"
                SetCellValueSafe ws.Cells(rowNo, "F"), "-"
                SetCellValueSafe ws.Cells(rowNo, "G"), "-"
                SetCellValueSafe ws.Cells(rowNo, "H"), "-"
            Else
                If CStr(GetCellValueSafe(ws.Cells(rowNo, "C"))) = "-" Then ClearCellSafe ws.Cells(rowNo, "C")
                If CStr(GetCellValueSafe(ws.Cells(rowNo, "E"))) = "-" Then ClearCellSafe ws.Cells(rowNo, "E")
                If CStr(GetCellValueSafe(ws.Cells(rowNo, "G"))) = "-" Then ClearCellSafe ws.Cells(rowNo, "G")
                If CStr(GetCellValueSafe(ws.Cells(rowNo, "H"))) = "-" Then ClearCellSafe ws.Cells(rowNo, "H")
                SetCellFormulaSafe ws.Cells(rowNo, "F"), "=C" & rowNo
            End If
        Else
            ClearUnusedDayBlock ws, rowNo
        End If
    Next dayIndex
End Sub

Private Sub ClearUnusedDayBlock(ByVal ws As Worksheet, ByVal rowNo As Long)
    ClearCellSafe ws.Cells(rowNo, "B")
    ClearCellSafe ws.Cells(rowNo + 1, "B")
    ClearCellSafe ws.Cells(rowNo + 3, "B")
    ClearCellSafe ws.Cells(rowNo + 4, "B")
    ClearCellSafe ws.Cells(rowNo, "C")
    ClearCellSafe ws.Cells(rowNo, "E")
    ClearCellSafe ws.Cells(rowNo, "F")
    ClearCellSafe ws.Cells(rowNo, "G")
    ClearCellSafe ws.Cells(rowNo, "H")
End Sub

Private Sub ClearCellSafe(ByVal targetCell As Range)
    If targetCell.MergeCells Then
        targetCell.MergeArea.ClearContents
    Else
        targetCell.ClearContents
    End If
End Sub

Private Sub SetCellValueSafe(ByVal targetCell As Range, ByVal newValue As Variant)
    If targetCell.MergeCells Then
        targetCell.MergeArea.Cells(1, 1).Value = newValue
    Else
        targetCell.Value = newValue
    End If
End Sub

Private Sub SetCellFormulaSafe(ByVal targetCell As Range, ByVal formulaText As String)
    If targetCell.MergeCells Then
        targetCell.MergeArea.Cells(1, 1).Formula = formulaText
    Else
        targetCell.Formula = formulaText
    End If
End Sub

Private Function GetCellValueSafe(ByVal targetCell As Range) As Variant
    If targetCell.MergeCells Then
        GetCellValueSafe = targetCell.MergeArea.Cells(1, 1).Value
    Else
        GetCellValueSafe = targetCell.Value
    End If
End Function

Private Function GetReportYear(ByVal reportMonth As Long) As Long
    If reportMonth >= 1 And reportMonth <= 3 Then
        GetReportYear = FISCAL_START_YEAR + 1
    Else
        GetReportYear = FISCAL_START_YEAR
    End If
End Function

Private Function GetOutputFolderPath() As String
    Dim p As String

    p = ThisWorkbook.Path
    If p = "" Then p = CurDir$
    If LCase$(Left$(p, 4)) = "http" Then p = CurDir$

    GetOutputFolderPath = p
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Private Function FileExists(ByVal fullPath As String) As Boolean
    FileExists = (Len(Dir$(fullPath, vbNormal)) > 0)
End Function

Private Function DeleteIfExists(ByVal fullPath As String) As Boolean
    On Error GoTo Fail

    If FileExists(fullPath) Then
        SetAttr fullPath, vbNormal
        Kill fullPath
    End If

    DeleteIfExists = True
    Exit Function

Fail:
    DeleteIfExists = False
End Function

Private Function IsWeekend(ByVal d As Date) As Boolean
    IsWeekend = (Weekday(d, vbMonday) >= 5)
End Function

Private Function PathsEqual(ByVal path1 As String, ByVal path2 As String) As Boolean
    PathsEqual = (StrComp(NormalizePath(path1), NormalizePath(path2), vbTextCompare) = 0)
End Function

Private Function NormalizePath(ByVal p As String) As String
    Dim s As String
    s = Replace(p, "/", "\\")
    Do While InStr(s, "\\\\") > 0
        s = Replace(s, "\\\\", "\\")
    Loop
    NormalizePath = s
End Function
