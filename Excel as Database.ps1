#This script will connect an excel as database and will help in retrieving the data

$path = "K:\CLE06\TEAM_RPA\RPA\Booking QV\CLS Mapping Exercise\CLS_Booked_Loan 2019-01-02.xlsx"
$connection = New-Object System.Data.OleDb.OleDbConnection
$connectstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$path;Extended Properties='Excel 12.0 Xml;HDR=YES';"
$connection.ConnectionString = $connectstring

#$connection.GetSchema()
#$connection.GetSchema("tables")
#$connection.GetSchema("columns")

$cmdObject = New-Object System.Data.OleDb.OleDbCommand
$cmdObjectAdapter = New-Object System.Data.OleDb.OleDbDataAdapter
$cmdObjectDataTables = New-Object System.Data.DataTable

$cmdObject.Connection = $connection
#$cmdObject.CommandText = "Select * from [Vendor$] where [Age]=26"
#$cmdObject.CommandText = "Select * from [Vendor$] where between [DataRange]= '1990-01-01 12:00:00' and '2001-01-01 12:00:00'"
#$cmdObject.CommandText = "Select * from [Vendor$] where [DataRange] > 1995-01-01"
#$cmdObject.CommandText = "Select * from [Vendor$] where between [DataRange] = #1990-01-01# and #1995-01-01#"
#$cmdObject.CommandText = "Select * from [Vendor$] (CASE Name WHEN 'Mike' THEN '1' WHEN 'Miller' THEN '2'ELSE Name END)"


$cmdObject.CommandText = "Select * from [Criteria$]"

$cmdObjectAdapter.SelectCommand = $cmdObject
$cmdObjectAdapter.Fill($cmdObjectDataTables)
$cmdObjectDataTables

$cmdObject.Dispose()
$connection.Close()
$connection.Dispose()