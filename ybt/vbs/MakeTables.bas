Attribute VB_Name = "MakeTables"
Sub MakeReviewWords()
'
' ���������ʱ� ��
'
Sheets("�����ʱ�").Select
Rem ɾ��ԭ������
Dim pos, rows
rows = Sheets("�����ʱ�").UsedRange.rows.Count
clos = Sheets("�����ʱ�").UsedRange.Columns.Count
Sheets("�´ʱ�").Range(Sheets("�´ʱ�").Cells(3, 1), Sheets("�´ʱ�").Cells(rows, clos)).Delete
  
Dim i
i = 3
For Each st In Worksheets
    If st.Name <> ActiveSheet.Name And st.ListObjects.Count And st.Cells(1, 2) >= 0 And st.Cells(1, 2) <= 0.2 Then
        Rem ��ȡ����
        rows = st.UsedRange.rows.Count
        Rem ��ʼѭ��
        For pos = 3 To rows
            If st.Cells(pos, 3) >= Sheets("�����ʱ�").Cells(1, 2) And st.Cells(pos, 3) <= Sheets("�����ʱ�").Cells(1, 3) Then
                st.Cells(pos, 3).EntireRow.Copy
                Sheets("�����ʱ�").Select
                Sheets("�����ʱ�").Cells(i, 1).Select
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
' �����´ʱ� ��
'
Rem ɾ��ԭ������
Sheets("�´ʱ�").Select
Dim pos, rows
rows = Sheets("�´ʱ�").UsedRange.rows.Count
clos = Sheets("�´ʱ�").UsedRange.Columns.Count
Sheets("�´ʱ�").Range(Sheets("�´ʱ�").Cells(3, 1), Sheets("�´ʱ�").Cells(rows, clos)).Delete

Dim i
i = 3
For Each st In Worksheets
    If st.Name <> ActiveSheet.Name And st.ListObjects.Count And st.Cells(1, 2) > 0.2 Then
        Rem ��ȡ����
        rows = st.UsedRange.rows.Count
        Rem ��ʼѭ��
        For pos = 3 To rows
            If st.Cells(pos, 3) >= Sheets("�´ʱ�").Cells(1, 2) And st.Cells(pos, 3) <= Sheets("�´ʱ�").Cells(1, 3) Then
                st.Cells(pos, 3).EntireRow.Copy
                Sheets("�´ʱ�").Select
                Sheets("�´ʱ�").Cells(i, 1).Select
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
' ͳ������ ��
'
Dim rows As Integer, clos  As Integer
rows = ActiveSheet.UsedRange.rows.Count
clos = ActiveSheet.UsedRange.Columns.Count
MsgBox rows
MsgBox clos
End Sub
