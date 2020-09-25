Attribute VB_Name = "BatchSort"

Sub sortAllByWord()
'
' sortAllByWord ��
'
Application.ScreenUpdating = False
For Each st In Worksheets
    If st.ListObjects.Count > 0 Then
        st.Activate
        Name = st.ListObjects(1).Name
        rangeName = Name + "[[#All],[word]]"
        st.ListObjects(1).Sort.SortFields.Clear
        st.ListObjects(1).Sort.SortFields.Add Key _
        :=Range(rangeName), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
        With st.ListObjects(1).Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
Next
Application.ScreenUpdating = True
End Sub


Sub sortAllByDate()
'
' sortAllByDate ��
'
Application.ScreenUpdating = False
For Each st In Worksheets
    If st.ListObjects.Count > 0 Then
        st.Activate
        Name = st.ListObjects(1).Name
        rangeName = Name + "[[#All],[���һ�����ǵ�����]]"
        st.ListObjects(1).Sort.SortFields.Clear
        st.ListObjects(1).Sort.SortFields.Add Key _
        :=Range(rangeName), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
        With st.ListObjects(1).Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If
Next
Application.ScreenUpdating = True
End Sub



