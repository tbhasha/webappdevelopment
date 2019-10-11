' Sub Macro2()
' '
' ' Macro2 Macro


' Set objExcel = CreateObject(Excel.Application)

    ' Range("A2:A15").Select
    ' Application.AddCustomList ListArray:=Array("David", "Lang")
    ' ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ' ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("A2:A15"), _
        ' SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:="David,Lang", _
        ' DataOption:=xlSortNormal
    ' With ActiveWorkbook.Worksheets("Sheet1").Sort
        ' .SetRange Range("A2:A15")
        ' .Header = xlGuess
        ' ' .MatchCase = False
        ' ' .Orientation = xlTopToBottom
        ' ' .SortMethod = xlPinYin
        ' .Apply
    ' End With
' End Sub

'Dim Excelpath = ‪"U:\VBSort TestData.xlsx"
Call ExcelSort()

Sub ExcelSort()
Const xlAscending = 1
Const xlDescending = 2
Const xlNo = 1

Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = 0
Set objWorkbook = objExcel.Workbooks.open("U:\VBSortTestData.xlsx")
objExcel.Application.Visible = True
Set objWorksheet = objWorkbook.worksheets(1)
objWorksheet.Activate

Set Objsheet2 = objWorkbook.worksheets(2)

' Set objRange = objWorksheet.UsedRange
' Set objRange2 = objExcel.Range("B1")
' objRange.Sort objRange2,xlAscending,,,,,xlNo
call CustomSorting()

Sub CustomSorting()
Dim r As Range
Dim rng As Range
Set r=Range("B10", Range("D" & Rows.Count).End(xlUp))
Set rng=objsheet2.Range("A2", Objsheet2.Range("A2").End(xlDown))
Application.AddCustomList rng
r.Sort key1:=[B10], order1:=1, ordercustom:=Application.CustomListCount + 1, , ,
Application.DeleteCustomList Application.CustomListCount
MsgBox " End of Custom Sort Function"

End Sub
'Sleep(2000)
'Range("A2:A15").Select
' Set objRange = objWorksheet.Range("A2:A15")
    ' objExcel.AddCustomList Array("David", "Lang")
    ' 'ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ' objWorksheet.Sort.SortFields.Clear
    ' 'objWorksheet.Sort.SortFields.Add Key=Range("A2:A15"), SortOn=xlSortOnValues, Order=xlAscending, CustomOrder="David,Lang", DataOption=xlSortNormal
    ' objWorksheet.Sort.SortFields.Add Key=objRange2, SortOn=xlSortOnValues, Order=xlAscending, CustomOrder="David,Lang", DataOption=xlSortNormal
    ' With objWorksheet.Sort
        ' .SetRange objRange2
        ' .Header = xlGuess
        ' .MatchCase = False
        ' .Orientation = xlTopToBottom
        ' .SortMethod = xlPinYin
        ' .Apply
    ' End With
	
objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit

MsgBox "Finished"

End Sub