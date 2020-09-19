Attribute VB_Name = "BatchFormat"
Sub ChangeTimeFormatBatch()
Attribute ChangeTimeFormatBatch.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �����޸�ʱ���ʽ ��
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
' ��������ȫ���ʽ ��
'
    n = Worksheets.Count
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
     MsgBox n
    For i = 1 To n
        Worksheets(i).Activate
        Cells.Select
        Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
 
Sub RestoreLogSheets()
'
' �ָ���־�� ��
'
    Sheets("��������־").Select
    Cells.Select
    Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
    Columns("A:A").Select
    Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
    Sheets("���и�ϰ�򿨱�").Select
    Cells.Select
    Selection.NumberFormatLocal = "G/ͨ�ø�ʽ"
    Columns("A:A").Select
    Selection.NumberFormatLocal = "[$-x-sysdate]dddd, mmmm dd, yyyy"
End Sub



