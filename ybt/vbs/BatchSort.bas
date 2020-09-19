Attribute VB_Name = "BatchSort"

Sub sortWord()
Attribute sortWord.VB_ProcData.VB_Invoke_Func = " \n14"
'
' sortWord 宏
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
    Sheets("Lv3L2").Select
    ActiveWorkbook.Worksheets("Lv3L2").ListObjects("Lv3L2T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L2").ListObjects("Lv3L2T1").Sort.SortFields.Add Key _
        :=Range("Lv3L2T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L2").ListObjects("Lv3L2T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L3").Select
    ActiveWorkbook.Worksheets("Lv3L3").ListObjects("Lv3L3T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L3").ListObjects("Lv3L3T1").Sort.SortFields.Add Key _
        :=Range("Lv3L3T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L3").ListObjects("Lv3L3T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L4").Select
    ActiveWorkbook.Worksheets("Lv3L4").ListObjects("Lv3L4T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L4").ListObjects("Lv3L4T1").Sort.SortFields.Add Key _
        :=Range("Lv3L4T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L4").ListObjects("Lv3L4T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L5").Select
    ActiveWorkbook.Worksheets("Lv3L5").ListObjects("Lv3L5T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L5").ListObjects("Lv3L5T1").Sort.SortFields.Add Key _
        :=Range("Lv3L5T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L5").ListObjects("Lv3L5T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L6").Select
    ActiveWorkbook.Worksheets("Lv3L6").ListObjects("Lv3L6T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L6").ListObjects("Lv3L6T1").Sort.SortFields.Add Key _
        :=Range("Lv3L6T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L6").ListObjects("Lv3L6T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L7").Select
    ActiveWorkbook.Worksheets("Lv3L7").ListObjects("Lv3L7T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L7").ListObjects("Lv3L7T1").Sort.SortFields.Add Key _
        :=Range("Lv3L7T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L7").ListObjects("Lv3L7T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L8").Select
    ActiveWorkbook.Worksheets("Lv3L8").ListObjects("Lv3L8T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L8").ListObjects("Lv3L8T1").Sort.SortFields.Add Key _
        :=Range("Lv3L8T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L8").ListObjects("Lv3L8T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L9").Select
    ActiveWorkbook.Worksheets("Lv3L9").ListObjects("Lv3L9T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L9").ListObjects("Lv3L9T1").Sort.SortFields.Add Key _
        :=Range("Lv3L9T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L9").ListObjects("Lv3L9T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L10").Select
    ActiveWorkbook.Worksheets("Lv3L10").ListObjects("Lv3L10T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L10").ListObjects("Lv3L10T1").Sort.SortFields.Add Key _
        :=Range("Lv3L10T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L10").ListObjects("Lv3L10T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L1").Select
    ActiveWorkbook.Worksheets("Lv4L1").ListObjects("Lv4L1T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L1").ListObjects("Lv4L1T1").Sort.SortFields.Add Key _
        :=Range("Lv4L1T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L1").ListObjects("Lv4L1T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L2").Select
    ActiveWorkbook.Worksheets("Lv4L2").ListObjects("Lv4L2T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L2").ListObjects("Lv4L2T1").Sort.SortFields.Add Key _
        :=Range("Lv4L2T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L2").ListObjects("Lv4L2T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L3").Select
    ActiveWorkbook.Worksheets("Lv4L3").ListObjects("Lv4L3T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L3").ListObjects("Lv4L3T1").Sort.SortFields.Add Key _
        :=Range("Lv4L3T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L3").ListObjects("Lv4L3T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L4").Select
    ActiveWorkbook.Worksheets("Lv4L4").ListObjects("Lv4L4T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L4").ListObjects("Lv4L4T1").Sort.SortFields.Add Key _
        :=Range("Lv4L4T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L4").ListObjects("Lv4L4T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L5").Select
    ActiveWorkbook.Worksheets("Lv4L5").ListObjects("Lv4L5T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L5").ListObjects("Lv4L5T1").Sort.SortFields.Add Key _
        :=Range("Lv4L5T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L5").ListObjects("Lv4L5T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L6").Select
    ActiveWorkbook.Worksheets("Lv4L6").ListObjects("Lv4L6T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L6").ListObjects("Lv4L6T1").Sort.SortFields.Add Key _
        :=Range("Lv4L6T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L6").ListObjects("Lv4L6T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L7").Select
    ActiveWorkbook.Worksheets("Lv4L7").ListObjects("Lv4L7T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L7").ListObjects("Lv4L7T1").Sort.SortFields.Add Key _
        :=Range("Lv4L7T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L7").ListObjects("Lv4L7T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L8").Select
    ActiveWorkbook.Worksheets("Lv4L8").ListObjects("Lv4L8T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L8").ListObjects("Lv4L8T1").Sort.SortFields.Add Key _
        :=Range("Lv4L8T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L8").ListObjects("Lv4L8T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L9").Select
    ActiveWorkbook.Worksheets("Lv4L9").ListObjects("Lv4L9T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L9").ListObjects("Lv4L9T1").Sort.SortFields.Add Key _
        :=Range("Lv4L9T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L9").ListObjects("Lv4L9T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L10").Select
    ActiveWorkbook.Worksheets("Lv4L10").ListObjects("Lv4L10T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L10").ListObjects("Lv4L10T1").Sort.SortFields.Add Key _
        :=Range("Lv4L10T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L10").ListObjects("Lv4L10T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv5L1").Select
    ActiveWorkbook.Worksheets("Lv5L1").ListObjects("Lv5L1T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv5L1").ListObjects("Lv5L1T1").Sort.SortFields.Add Key _
        :=Range("Lv5L1T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv5L1").ListObjects("Lv5L1T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv5L2").Select
    ActiveWorkbook.Worksheets("Lv5L2").ListObjects("Lv5L2T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv5L2").ListObjects("Lv5L2T1").Sort.SortFields.Add Key _
        :=Range("Lv5L2T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv5L2").ListObjects("Lv5L2T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv5L3").Select
    ActiveWorkbook.Worksheets("Lv5L3").ListObjects("Lv5L3T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv5L3").ListObjects("Lv5L3T1").Sort.SortFields.Add Key _
        :=Range("Lv5L3T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv5L3").ListObjects("Lv5L3T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv5L4").Select
    ActiveWorkbook.Worksheets("Lv5L4").ListObjects("Lv5L4T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv5L4").ListObjects("Lv5L4T1").Sort.SortFields.Add Key _
        :=Range("Lv5L4T1[[#All],[word]]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv5L4").ListObjects("Lv5L4T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub sortDate()
Attribute sortDate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' sortDate 宏
'

'
    ActiveWorkbook.Worksheets("Lv3L1").ListObjects("Lv3L1T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L1").ListObjects("Lv3L1T1").Sort.SortFields.Add Key _
        :=Range("Lv3L1T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L1").ListObjects("Lv3L1T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L2").Select
    ActiveWorkbook.Worksheets("Lv3L2").ListObjects("Lv3L2T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L2").ListObjects("Lv3L2T1").Sort.SortFields.Add Key _
        :=Range("Lv3L2T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L2").ListObjects("Lv3L2T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L3").Select
    ActiveWorkbook.Worksheets("Lv3L3").ListObjects("Lv3L3T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L3").ListObjects("Lv3L3T1").Sort.SortFields.Add Key _
        :=Range("Lv3L3T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L3").ListObjects("Lv3L3T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L4").Select
    ActiveWorkbook.Worksheets("Lv3L4").ListObjects("Lv3L4T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L4").ListObjects("Lv3L4T1").Sort.SortFields.Add Key _
        :=Range("Lv3L4T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L4").ListObjects("Lv3L4T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L5").Select
    ActiveWorkbook.Worksheets("Lv3L5").ListObjects("Lv3L5T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L5").ListObjects("Lv3L5T1").Sort.SortFields.Add Key _
        :=Range("Lv3L5T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L5").ListObjects("Lv3L5T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L6").Select
    ActiveWorkbook.Worksheets("Lv3L6").ListObjects("Lv3L6T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L6").ListObjects("Lv3L6T1").Sort.SortFields.Add Key _
        :=Range("Lv3L6T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L6").ListObjects("Lv3L6T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L7").Select
    ActiveWorkbook.Worksheets("Lv3L7").ListObjects("Lv3L7T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L7").ListObjects("Lv3L7T1").Sort.SortFields.Add Key _
        :=Range("Lv3L7T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L7").ListObjects("Lv3L7T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L8").Select
    ActiveWorkbook.Worksheets("Lv3L8").ListObjects("Lv3L8T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L8").ListObjects("Lv3L8T1").Sort.SortFields.Add Key _
        :=Range("Lv3L8T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L8").ListObjects("Lv3L8T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L9").Select
    ActiveWorkbook.Worksheets("Lv3L9").ListObjects("Lv3L9T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L9").ListObjects("Lv3L9T1").Sort.SortFields.Add Key _
        :=Range("Lv3L9T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L9").ListObjects("Lv3L9T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv3L10").Select
    ActiveWorkbook.Worksheets("Lv3L10").ListObjects("Lv3L10T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv3L10").ListObjects("Lv3L10T1").Sort.SortFields.Add Key _
        :=Range("Lv3L10T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv3L10").ListObjects("Lv3L10T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L1").Select
    ActiveWorkbook.Worksheets("Lv4L1").ListObjects("Lv4L1T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L1").ListObjects("Lv4L1T1").Sort.SortFields.Add Key _
        :=Range("Lv4L1T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L1").ListObjects("Lv4L1T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L2").Select
    ActiveWorkbook.Worksheets("Lv4L2").ListObjects("Lv4L2T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L2").ListObjects("Lv4L2T1").Sort.SortFields.Add Key _
        :=Range("Lv4L2T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L2").ListObjects("Lv4L2T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L3").Select
    ActiveWorkbook.Worksheets("Lv4L3").ListObjects("Lv4L3T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L3").ListObjects("Lv4L3T1").Sort.SortFields.Add Key _
        :=Range("Lv4L3T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L3").ListObjects("Lv4L3T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L4").Select
    ActiveWorkbook.Worksheets("Lv4L4").ListObjects("Lv4L4T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L4").ListObjects("Lv4L4T1").Sort.SortFields.Add Key _
        :=Range("Lv4L4T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L4").ListObjects("Lv4L4T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L5").Select
    ActiveWorkbook.Worksheets("Lv4L5").ListObjects("Lv4L5T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L5").ListObjects("Lv4L5T1").Sort.SortFields.Add Key _
        :=Range("Lv4L5T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L5").ListObjects("Lv4L5T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L6").Select
    ActiveWorkbook.Worksheets("Lv4L6").ListObjects("Lv4L6T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L6").ListObjects("Lv4L6T1").Sort.SortFields.Add Key _
        :=Range("Lv4L6T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L6").ListObjects("Lv4L6T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L7").Select
    ActiveWorkbook.Worksheets("Lv4L7").ListObjects("Lv4L7T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L7").ListObjects("Lv4L7T1").Sort.SortFields.Add Key _
        :=Range("Lv4L7T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L7").ListObjects("Lv4L7T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L8").Select
    ActiveWorkbook.Worksheets("Lv4L8").ListObjects("Lv4L8T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L8").ListObjects("Lv4L8T1").Sort.SortFields.Add Key _
        :=Range("Lv4L8T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L8").ListObjects("Lv4L8T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L9").Select
    ActiveWorkbook.Worksheets("Lv4L9").ListObjects("Lv4L9T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L9").ListObjects("Lv4L9T1").Sort.SortFields.Add Key _
        :=Range("Lv4L9T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L9").ListObjects("Lv4L9T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv4L10").Select
    ActiveWorkbook.Worksheets("Lv4L10").ListObjects("Lv4L10T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv4L10").ListObjects("Lv4L10T1").Sort.SortFields.Add Key _
        :=Range("Lv4L10T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv4L10").ListObjects("Lv4L10T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv5L1").Select
    ActiveWorkbook.Worksheets("Lv5L1").ListObjects("Lv5L1T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv5L1").ListObjects("Lv5L1T1").Sort.SortFields.Add Key _
        :=Range("Lv5L1T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv5L1").ListObjects("Lv5L1T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv5L2").Select
    ActiveWorkbook.Worksheets("Lv5L2").ListObjects("Lv5L2T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv5L2").ListObjects("Lv5L2T1").Sort.SortFields.Add Key _
        :=Range("Lv5L2T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv5L2").ListObjects("Lv5L2T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv5L3").Select
    ActiveWorkbook.Worksheets("Lv5L3").ListObjects("Lv5L3T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv5L3").ListObjects("Lv5L3T1").Sort.SortFields.Add Key _
        :=Range("Lv5L3T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv5L3").ListObjects("Lv5L3T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Lv5L4").Select
    ActiveWorkbook.Worksheets("Lv5L4").ListObjects("Lv5L4T1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Lv5L4").ListObjects("Lv5L4T1").Sort.SortFields.Add Key _
        :=Range("Lv5L4T1[[#All],[最后一次忘记的日期]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Lv5L4").ListObjects("Lv5L4T1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
