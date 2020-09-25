Attribute VB_Name = "BatchFormat"
Sub RestoreAllFormatSheets()
'
' �ָ����б�ĸ�ʽ ��
'
    Rem ����ȫ����ʽ�������õ�����Ϊ���ڣ�ǰ���ű�����Ϊ����
    n = Worksheets.Count
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    MsgBox n
    For Each st In Worksheets
        st.Activate
        st.Cells.Select
        Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
        If st.Name <> "����˵��" And st.Name <> "��������־" And st.Name <> "���и�ϰ�򿨱�" Then
            Columns("C:C").Select
            Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
        End If
        If st.Name = "����˵��" Or st.Name = "��������־" Or st.Name = "���и�ϰ�򿨱�" Then
            Columns("A:A").Select
            Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
        End If
        If st.Name = "�����ʱ�" Or st.Name = "�´ʱ�" Then
            rows("1:1").Select
            Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
        End If
        st.Cells.Select
        Selection.Columns.AutoFit
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
