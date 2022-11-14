* Change these as necessary.

lcSource     = 'C:\MyData'
	&& Folder containing VFP free tables or name and path of a DBC
lcConnString = 'server=MyServer;database=MyDatabase;trusted_connection=true'
	&& Connection string for SQL Server database
lcTable      = 'MyTable'
	&& Table to bulk copy

* Delete existing records.

lnHandle = sqlstringconnect('driver=SQL Server;' + lcConnString)
if lnHandle > 0
    lnResult = sqlexec(lnHandle, 'TRUNCATE TABLE [' + lcTable + ']')
    if lnResult < 0
* handle error
	endif
    sqldisconnect(lnHandle)
endif

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
