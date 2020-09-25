Attribute VB_Name = "BatchFormat"
Sub RestoreAllFormatSheets()
'
' 恢复所有表的格式 宏
'
    Rem 消除全部格式，并设置第三列为日期，前三张表首列为日期
    n = Worksheets.Count
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    MsgBox n
    For Each st In Worksheets
        st.Activate
        st.Cells.Select
        Selection.NumberFormatLocal = "G/通用格式"
        If st.Name <> "总述说明" And st.Name <> "背单词日志" And st.Name <> "背诵复习打卡表" Then
            Columns("C:C").Select
            Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
        End If
        If st.Name = "总述说明" Or st.Name = "背单词日志" Or st.Name = "背诵复习打卡表" Then
            Columns("A:A").Select
            Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
        End If
        If st.Name = "易忘词表" Or st.Name = "新词表" Then
            rows("1:1").Select
            Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
        End If
        st.Cells.Select
        Selection.Columns.AutoFit
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
