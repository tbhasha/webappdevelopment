$ExcelConnection= New-Object -com "ADODB.Connection" 
$ExcelFile="c:\temp\VendorList.xlsx" 
$ExcelConnection.Open("Provider=Microsoft.ACE.OLEDB.12.0;` 
Data Source=$ExcelFile;Extended Properties=Excel 12.0;")

$strQuery="Select * from [Vendors$]" 
$ExcelRecordSet=$ExcelConnection.Execute($strQuery)

do { 
Write-Host "EXEC sp_InsertVendors '" $ExcelRecordSet.Fields.Item("Vendor Code").Value "'" 
$ExcelRecordSet.MoveNext()} 
Until ($ExcelRecordSet.EOF)

$ExcelConnection.Close()