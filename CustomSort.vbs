'**************************************************************************************************
'	Description		:  CustomSort VBScript is used to applying custom sort for excel 
'					   input provided
'   Change History 	:  	
'**************************************************************************************************
'**************************************************************************************************
'Declare required variables
'**************************************************************************************************
Dim objExcel
Dim objWorkbook
Dim objWorksheet
Dim inputFilePath
Dim isError
Dim xlSortNormal,xlSortOnValues,xlAscending,xlTopToBottom,xlPinYin,xlYes

xlSortNormal = 0
xlSortOnValues = 0
xlAscending = 1
xlTopToBottom = 1	
xlPinYin = 1
xlYes = 1
isError = True
'**************************************************************************************************
' Begin Error Handling
'**************************************************************************************************
On Error Resume Next


'inputFilePath = WScript.Arguments(1) 
inputFilePath = "U:\VB Scripts\PAYOFF_REPORT_0322 01 55 PM.xlsm"

'Open the excel file
Set objExcel = CreateObject("Excel.Application") 
'msgbox "excel object created"
Set objWorkbook = objExcel.Workbooks.Open(inputFilePath)
Set objWorksheet = objWorkbook.Worksheets("PAYOFF_REPORT")
'msgbox "excel open"
'Sort by custom array
objWorksheet.Sort.SortFields.Clear
'objWorksheet.Sort.SortFields.Add key:=Range("C2"), SortOn:=xlSortOnValues, Order:=xlAscending,CustomOrder:="AUTOU", DataOption:=xlSortNormal
objWorksheet.Sort.SortFields.Clear
objWorksheet.Sort.SortFields.Add objWorksheet.Range("C:C"), xlSortOnValues, xlAscending,"BOATU,AUTOU", xlSortNormal
With objWorksheet.Sort
.SetRange objWorksheet.Range("A:SX")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'msgbox "sort applied"
'Save Excel workbook with same name
objWorkbook.save 
'msgbox "save"
objExcel.quit 


If Err.Number <> 0 Then		
		isError = "Err Description " & Err.Description & " Err No " & Err.Number
		Err.Clear		
End If

WScript.StdOut.WriteLine (isError)