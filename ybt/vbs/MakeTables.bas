Attribute VB_Name = "MakeTables"
Sub MakeReviewWords()
'
' ���������ʱ� ��
'
Application.ScreenUpdating = False
Sheets("�����ʱ�").Select
Rem ɾ��ԭ������
Dim pos, rows
rows = Sheets("�����ʱ�").UsedRange.rows.Count
clos = Sheets("�����ʱ�").UsedRange.Columns.Count
Sheets("�����ʱ�").Range(Sheets("�����ʱ�").Cells(3, 1), Sheets("�����ʱ�").Cells(rows, clos)).Delete
  
Dim i
i = 3
For Each st In Worksheets
    If st.Name <> "����˵��" And st.Name <> "��������־" And st.Name <> "���и�ϰ�򿨱�" And st.Name <> "�����ʱ�" And st.Name <> "�´ʱ�" And st.ListObjects.Count > 0 And st.Cells(1, 2) > 0 And st.Cells(1, 2) <= 0.2 Then
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
Application.ScreenUpdating = True
End Sub
Sub MakeNewListWords()
'
' �����´ʱ� ��
'
Application.ScreenUpdating = False
Rem ɾ��ԭ������
Sheets("�´ʱ�").Select
Dim pos, rows
rows = Sheets("�´ʱ�").UsedRange.rows.Count
clos = Sheets("�´ʱ�").UsedRange.Columns.Count
Sheets("�´ʱ�").Range(Sheets("�´ʱ�").Cells(3, 1), Sheets("�´ʱ�").Cells(rows, clos)).Delete

Dim i
i = 3
For Each st In Worksheets
     If st.Name <> "����˵��" And st.Name <> "��������־" And st.Name <> "���и�ϰ�򿨱�" And st.Name <> "�����ʱ�" And st.Name <> "�´ʱ�" And st.ListObjects.Count > 0 And st.Cells(1, 2) > 0.2 Then
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
Application.ScreenUpdating = False
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

Sub MakeAllWords()
'
' ����ȫ�ʱ� ��
'
Sheets("�ʻ㵼����ʱ��").Select
Rem ɾ��ԭ������
Dim pos, rows
rows = Sheets("�ʻ㵼����ʱ��").UsedRange.rows.Count
clos = Sheets("�ʻ㵼����ʱ��").UsedRange.Columns.Count
Sheets("�ʻ㵼����ʱ��").Range(Sheets("�ʻ㵼����ʱ��").Cells(3, 1), Sheets("�ʻ㵼����ʱ��").Cells(rows, clos)).Delete
  
Dim i
i = 3
For Each st In Worksheets
     If st.Name <> "����˵��" And st.Name <> "��������־" And st.Name <> "���и�ϰ�򿨱�" And st.Name <> "�����ʱ�" And st.Name <> "�´ʱ�" And IsNumeric(st.Cells(1, 1)) And st.Cells(1, 1) >= 1 Then
        Rem ��ȡ����
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
        Rem ����ճ����Ч��Χ
        st.Range(st.Cells(3, 1), st.Cells(rows, clos)).Copy
        Sheets("�ʻ㵼����ʱ��").Select
        Sheets("�ʻ㵼����ʱ��").Cells(i, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
        i = i + rows - 2
    End If
Next
End Sub
