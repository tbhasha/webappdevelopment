''Converting Text to Columns with Delimiter as Pipe Symbol


Dim objExcel
Dim objWorkbook
Dim objWorksheet
Dim inputFilePath
Dim isError


isError = True
InputFilePath = WScript.Arguments(0)
OutputFileName = WScript.Arguments(1)
LogFileName = WScript.Arguments(2)
'InputFilePath = "C:\Users\pallath\Desktop\BOT CLS EXTRACT.csv"
'OutputFileName = "C:\Users\pallath\Desktop\Extraced Data.xlsx"

'**************************************************************************************************
' Begin Error Handling
'**************************************************************************************************
On Error Resume Next
Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(LogFileName,8,true)

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(InputFilePath)
Set objWorksheet = objWorkbook.Worksheets(1)
Dim objRange
'objRange = objWorksheet.Range("A:A").Select

objExcel.visible = True
objExcel.DisplayAlerts = False
LastRow=objWorksheet.UsedRange.Rows.Count

'Rows("1:1").Select
'ActiveWindow.FreezePanes = True
'objWorksheet.Columns("A:A").TextToColumns objWorksheet.range("A1"), 1, , , , , , , True, "|"
'objWorksheet.Range("A2",Range("A2").End(xlDown)).TextToColumns objWorksheet.range("A2"), 1, , , , , , , True, "|"
objWorksheet.Range("A2","A"&LastRow).TextToColumns objWorksheet.Range("A2"), 1, , , , , , , True, "|"

objWorkbook.SaveAs OutputFileName,51
objWorkbook.Close()
objExcel.Quit

If Err.Number <> 0 Then		
		isError = "Err Description " & Err.Description & " Err No " & Err.Number
		objFileToWrite.WriteLine("("&Now()&") "&"Error	" & "ConvertTexttoColumns" & Err.Description )
		vErrorDescription = isError
		Err.Clear		
End If
'**************************************************************************************************
' End of Error Handling
'**************************************************************************************************

WScript.StdOut.WriteLine (isError)