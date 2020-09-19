Attribute VB_Name = "MakeTables"
Sub MakeReviewWords()
'
' 构造易忘词表 宏
'
Sheets("易忘词表").Select
Rem 删除原有数据
Dim pos, rows
rows = Sheets("易忘词表").UsedRange.rows.Count
clos = Sheets("易忘词表").UsedRange.Columns.Count
Sheets("新词表").Range(Sheets("新词表").Cells(3, 1), Sheets("新词表").Cells(rows, clos)).Delete
  
Dim i
i = 3
For Each st In Worksheets
    If st.Name <> ActiveSheet.Name And st.ListObjects.Count And st.Cells(1, 2) >= 0 And st.Cells(1, 2) <= 0.2 Then
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
End Sub
Sub MakeNewListWords()
'
' 构造新词表 宏
'
Rem 删除原有数据
Sheets("新词表").Select
Dim pos, rows
rows = Sheets("新词表").UsedRange.rows.Count
clos = Sheets("新词表").UsedRange.Columns.Count
Sheets("新词表").Range(Sheets("新词表").Cells(3, 1), Sheets("新词表").Cells(rows, clos)).Delete

Dim i
i = 3
For Each st In Worksheets
    If st.Name <> ActiveSheet.Name And st.ListObjects.Count And st.Cells(1, 2) > 0.2 Then
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
