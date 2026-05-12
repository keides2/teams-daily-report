VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const FIRST_BLOCK_ROW As Long = 21
Private Const BLOCK_HEIGHT As Long = 8
Private Const LAST_POSSIBLE_DAY As Long = 31

Private Sub Workbook_Open()
    RefreshAllReportSheets
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    RefreshAllReportSheets
End Sub

' 手動実行用
Public Sub RefreshAllReportSheets()
    Dim ws As Worksheet

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each ws In ThisWorkbook.Worksheets
        If IsTargetReportSheet(ws.Name) Then
            ApplyMonthlyLayout ws
        End If
    Next ws

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "ブック内シートの更新中にエラーが発生しました: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub ApplyMonthlyLayout(ByVal ws As Worksheet)
    Dim reportYear As Long
    Dim reportMonth As Long
    Dim lastDay As Long
    Dim dayIndex As Long
    Dim rowNo As Long
    Dim d As Date

    If Not TryParseYearMonthFromSheetName(ws.Name, reportYear, reportMonth) Then Exit Sub

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

Private Function IsTargetReportSheet(ByVal sheetName As String) As Boolean
    IsTargetReportSheet = (InStr(sheetName, "実施申請・報告書（") > 0 And InStr(sheetName, "年") > 0 And InStr(sheetName, "月") > 0)
End Function

Private Function TryParseYearMonthFromSheetName(ByVal sheetName As String, ByRef reportYear As Long, ByRef reportMonth As Long) As Boolean
    Dim pOpen As Long
    Dim pYear As Long
    Dim pMonth As Long
    Dim yearText As String
    Dim monthText As String

    pOpen = InStr(sheetName, "（")
    pYear = InStr(sheetName, "年")
    pMonth = InStr(sheetName, "月")

    If pOpen = 0 Or pYear = 0 Or pMonth = 0 Then Exit Function
    If Not (pOpen < pYear And pYear < pMonth) Then Exit Function

    yearText = Mid$(sheetName, pOpen + 1, pYear - pOpen - 1)
    monthText = Mid$(sheetName, pYear + 1, pMonth - pYear - 1)

    If Not IsNumeric(yearText) Then Exit Function
    If Not IsNumeric(monthText) Then Exit Function

    reportYear = CLng(yearText)
    reportMonth = CLng(monthText)

    If reportMonth < 1 Or reportMonth > 12 Then Exit Function

    TryParseYearMonthFromSheetName = True
End Function

Private Function IsWeekend(ByVal d As Date) As Boolean
    IsWeekend = (Weekday(d, vbMonday) >= 5)
End Function

