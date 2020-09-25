Attribute VB_Name = "Record"
Sub RecordSortWord()
'
' Â¼ÖÆµÄsortWord ºê
'
'
    ActiveWorkbook.Worksheets("Lv3L1").ListObjects("Lv3L1T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L1").ListObjects("Lv3L1T1").Sort.SortFields.Add Key _
        :=Range("Lv3L1T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L1").ListObjects("Lv3L1T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

