Attribute VB_Name = "ExportToFile"
Sub ConvertToCsv()
'
' 导出csv文件 宏
'
    Dim sh As Worksheet, p$
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = True
        .Show
        If .SelectedItems.Count = 0 Then Exit Sub
        p = .SelectedItems(1) & "\tywl_sh_"
        MsgBox p
    End With
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each sh In ThisWorkbook.Sheets
        sh.Copy
        With ActiveWorkbook
            .SaveAs Filename:=p & sh.Name & ".csv", FileFormat:=xlCSV
            .Close
        End With
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "ok"
End Sub
