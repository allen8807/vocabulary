Attribute VB_Name = "BatchFormat"
Sub ChangeTimeFormatBatch()
Attribute ChangeTimeFormatBatch.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 批量修改时间格式 宏
'
    n = Worksheets.Count
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
     MsgBox n
    For i = 1 To n
        Worksheets(i).Activate
        Columns("C:C").Select
        Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Sub RestoreTimeFormatBatch()
'
' 批量消除全表格式 宏
'
    n = Worksheets.Count
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
     MsgBox n
    For i = 1 To n
        Worksheets(i).Activate
        Cells.Select
        Selection.NumberFormatLocal = "G/通用格式"
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
 
Sub RestoreLogSheets()
'
' 恢复日志表 宏
'
    Sheets("背单词日志").Select
    Cells.Select
    Selection.NumberFormatLocal = "G/通用格式"
    Columns("A:A").Select
    Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
    Sheets("背诵复习打卡表").Select
    Cells.Select
    Selection.NumberFormatLocal = "G/通用格式"
    Columns("A:A").Select
    Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
End Sub



