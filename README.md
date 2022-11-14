# SQL Bulk Copy

## Introduction

The [VFP Upsizing Wizard](https://github.com/VFPX/UpsizingWizard) does a great job of upsizing a VFP database to SQL Server. However, after doing the initial upsizing, you may wish to re-upload the content of one or more tables without doing a complete re-upsize, such as if data entry into the VFP tables has continued after doing a test upsizing. This project is great for that.

This project uses the [.NET SqlBulkCopy class](https://learn.microsoft.com/en-us/dotnet/api/system.data.sqlclient.sqlbulkcopy) to upsize the content of a VFP table (free table or part of a DBC) to a SQL Server table. It's fast: it can upload about 15,000 records per second. It's also easy to use: just use [wwDotNetBridge](https://github.com/RickStrahl/wwDotnetBridge) to instantiate the SQLBulkCopy wrapper class that's part of this project, set the SourceConnectionString property to the  connection string for the VFP table (see below), set the DestinationConnectionString to the connection string for the SQL Server database, and call the LoadData method. Here's an example:

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
	loBulkCopy.LoadData(lcTable, '[' + lcTable + ']')
catch to loException
	messagebox('Error #' + transform(loException.ErrorNo) + ;
		' occurred in line ' + transform(loException.LineNo) + ;
		' of ' + loException.Procedure + ': ' + loException.Message, 16, ;
		'Upsize Error')
endtry
```
## Notes

* This example uses the VFP OLE DB provider (available from the [VFPX VFP 9 SP2 Hotfix 3 repository](https://github.com/VFPX/VFP9SP2Hotfix3) if not already installed). You can also use the VFP ODBC Driver by changing the connection string to something like "Driver=Microsoft Visual FoxPro Driver;SourceType=DBC;SourceDB=c:\myvfpdb.dbc". However, it's significantly slower than OLE DB (about half the speed) and doesn't support features added to VFP tables after version 6.

* DestinationConnectionString must be a valid [.NET SqlClient connection string](https://www.connectionstrings.com/microsoft-data-sqlclient/).

* The SQL Server database and table must exist and the table must have the same column names and data types as the VFP table (usually the case after upsizing).

* This project appends records to existing ones in the table, so if you want to completely replace the content of the table, use something like this before calling LoadData:

    ```foxpro
    lnHandle = sqlstringconnect('driver=SQL Server;' + lcConnString)
    if lnHandle > 0
        lnResult = sqlexec(lnHandle, 'TRUNCATE TABLE [' + lcTable + ']')
        if lnResult < 0
        * handle error
        endif
        sqldisconnect(lnHandle)
    endif
    ```

## Releases

### 2022-11-14

* Initial release