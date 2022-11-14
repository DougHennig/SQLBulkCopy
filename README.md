# SQL Bulk Copy

## Introduction

The [VFP Upsizing Wizard](https://github.com/VFPX/UpsizingWizard) does a great job of upsizing a VFP database to SQL Server. However, after doing the initial upsizing, you may wish to re-upload the content of one or more tables with doing a complete re-upsize. This project is great for that.

This project uses the .NET SqlBulkCopy class to upsize the content of a VFP table (free table or part of a DBC) to a SQL Server table. It's extremely fast: it can upload about 15,000 records per second. It's also really easy to use: just use wwDotNetBridge ***LINK to instantiate the SQLBulkCopy wrapper class that's part of this project, set the SourceConnectionString property to the  connection string for the VFP OLE DB provider to access the table, set the DestinationConnectionString to the connection strong for the SQL Server database, and call the LoadData method. Here's an example:

```foxpro
* Change these as necessary.

lcSource     = 'C:\MyData'
	&& Folder containing VFP free tables or name and path of a DBC
lcConnString = 'server=MyServer;database=MyDatabase;trusted_connection=true'
	&& Connection string for SQL Server database
lcTable      = 'MyTable'
	&& Table to bulk copy

* Bulk copy the VFP table content to SQL Server.

do wwDotNetBridge
loBridge = GetwwDotNetBridge()
loBridge.LoadAssembly('SQLBulkCopy.dll')
loBulkCopy = loBridge.CreateInstance('SQLBulkCopy.SQLBulkCopy')
loBulkCopy.SourceConnectionString      = 'provider=vfpoledb;' + ;
	'Data Source=' + lcSource
loBulkCopy.DestinationConnectionString = lcConnString
try
	loBulkCopy.LoadData(lcTable', '[' + lcTable + ']')
catch to loException
	messagebox('Error #' + transform(loException.ErrorNo) + ;
		' occurred in line ' + transform(loException.LineNo) + ;
		' of ' + loException.Procedure + ': ' + loException.Message, 16, ;
		'Upsize Error')
endtry
```
## Notes

* You can also use the VFP ODBC Driver by changing the connection string to something like "Driver=Microsoft Visual FoxPro Driver;SourceType=DBC;SourceDB=c:\myvfpdb.dbc". However, it's significantly slower than OLE DB (about half the speed) and doesn't support features added to VFP tables after version 6.

* DestinationConnectionString must be a valid [.NET SqlClient connection string](https://www.connectionstrings.com/microsoft-data-sqlclient/).

* The SQL Server database and table must exist and the table must have the same column names and data types as the VFP table (usually the case after upsizing).

## Releases

### 2022-11-14

* Initial release