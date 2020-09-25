Attribute VB_Name = "MakeTables"
Sub MakeReviewWords()
'
' 构造易忘词表 宏
'
Application.ScreenUpdating = False
Sheets("易忘词表").Select
Rem 删除原有数据
Dim pos, rows
rows = Sheets("易忘词表").UsedRange.rows.Count
clos = Sheets("易忘词表").UsedRange.Columns.Count
Sheets("易忘词表").Range(Sheets("易忘词表").Cells(3, 1), Sheets("易忘词表").Cells(rows, clos)).Delete
  
Dim i
i = 3
For Each st In Worksheets
    If st.Name <> "总述说明" And st.Name <> "背单词日志" And st.Name <> "背诵复习打卡表" And st.Name <> "易忘词表" And st.Name <> "新词表" And st.ListObjects.Count > 0 And st.Cells(1, 2) > 0 And st.Cells(1, 2) <= 0.2 Then
        Rem 获取行数
        rows = st.UsedRange.rows.Count
        Rem 开始循环
        For pos = 3 To rows
            If st.Cells(pos, 3) >= Sheets("易忘词表").Cells(1, 2) And st.Cells(pos, 3) <= Sheets("易忘词表").Cells(1, 3) Then
                st.Cells(pos, 3).EntireRow.Copy
                Sheets("易忘词表").Select
                Sheets("易忘词表").Cells(i, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                xlNone, SkipBlanks:=False, Transpose:=False
                i = i + 1
            End If
        Next
    End If
Next
Application.ScreenUpdating = True
End Sub
Sub MakeNewListWords()
'
' 构造新词表 宏
'
Application.ScreenUpdating = False
Rem 删除原有数据
Sheets("新词表").Select
Dim pos, rows
rows = Sheets("新词表").UsedRange.rows.Count
clos = Sheets("新词表").UsedRange.Columns.Count
Sheets("新词表").Range(Sheets("新词表").Cells(3, 1), Sheets("新词表").Cells(rows, clos)).Delete

Dim i
i = 3
For Each st In Worksheets
     If st.Name <> "总述说明" And st.Name <> "背单词日志" And st.Name <> "背诵复习打卡表" And st.Name <> "易忘词表" And st.Name <> "新词表" And st.ListObjects.Count > 0 And st.Cells(1, 2) > 0.2 Then
        Rem 获取行数
        rows = st.UsedRange.rows.Count
        Rem 开始循环
        For pos = 3 To rows
            If st.Cells(pos, 3) >= Sheets("新词表").Cells(1, 2) And st.Cells(pos, 3) <= Sheets("新词表").Cells(1, 3) Then
                st.Cells(pos, 3).EntireRow.Copy
                Sheets("新词表").Select
                Sheets("新词表").Cells(i, 1).Select
                Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                xlNone, SkipBlanks:=False, Transpose:=False
                i = i + 1
            End If
        Next
    End If
Next
Application.ScreenUpdating = False
End Sub

Sub StatsRC()
'
' 统计行列 宏
'
Dim rows As Integer, clos  As Integer
rows = ActiveSheet.UsedRange.rows.Count
clos = ActiveSheet.UsedRange.Columns.Count
MsgBox rows
MsgBox clos
End Sub

Sub MakeAllWords()
'
' 构造全词表 宏
'
Sheets("词汇导出临时表").Select
Rem 删除原有数据
Dim pos, rows
rows = Sheets("词汇导出临时表").UsedRange.rows.Count
clos = Sheets("词汇导出临时表").UsedRange.Columns.Count
Sheets("词汇导出临时表").Range(Sheets("词汇导出临时表").Cells(3, 1), Sheets("词汇导出临时表").Cells(rows, clos)).Delete
  
Dim i
i = 3
For Each st In Worksheets
     If st.Name <> "总述说明" And st.Name <> "背单词日志" And st.Name <> "背诵复习打卡表" And st.Name <> "易忘词表" And st.Name <> "新词表" And IsNumeric(st.Cells(1, 1)) And st.Cells(1, 1) >= 1 Then
        Rem 获取行数
        rows = st.Cells(2, 4).End(xlDown).Row
        clos = st.UsedRange.Columns.Count
        If i < 10 Then
            MsgBox rows
            MsgBox clos
        End If
        If rows > 180 Then
            MsgBox st.Name
            MsgBox rows
            MsgBox clos
        End If
        Rem 复制粘贴有效范围
        st.Range(st.Cells(3, 1), st.Cells(rows, clos)).Copy
        Sheets("词汇导出临时表").Select
        Sheets("词汇导出临时表").Cells(i, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        i = i + rows - 2
    End If
Next
End Sub
