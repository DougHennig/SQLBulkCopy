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
else
* handle error
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

* Create scripts to replace empty dates with null and strip spaces from
* character columns.

use (addbs(lcSource) + lcTable)
lnTotalFields = afields(laFields)
llFirst       = .T.
lcFixDateSQL  = ''
lcVarcharSQL  = ''
for lnFieldNo = 1 to lnTotalFields
	lcFieldName = laFields[lnFieldNo, 1]
	lcFieldType = laFields[lnFieldNo, 2]
	do case
		case lcFieldType $ 'DT'
			text to lcFixDateSQL noshow textmerge pretext 2
			update [<<lcTable>>] set [<<lcFieldName>>] = null where [<<lcFieldName>>] = '1899-12-30'
			
			endtext
		case lcFieldType = 'C' and llFirst
			text to lcVarcharSQL noshow textmerge pretext 2
			update [<<lcTable>>] set [<<lcFieldName>>] = RTRIM([<<lcFieldName>>])
			endtext
			llFirst = .F.
		case lcFieldType = 'C'
			text to lcVarcharSQL additive noshow textmerge pretext 2
			, [<<lcFieldName>>] = RTRIM([<<lcFieldName>>])
			endtext
	endcase
next
use

* Execute the scripts.

if not empty(lcFixDateSQL + lcVarcharSQL)
	lnHandle = sqlstringconnect('driver=SQL Server;' + lcConnString)
	if lnHandle > 0
		sqlsetprop(lnHandle, 'QueryTimeOut', 0)
		if not empty(lcFixDateSQL)
			lnResult = sqlexec(lnHandle, lcFixDateSQL)
			if lnResult < 0
			* handle error
			endif
		endif
		if not empty(lcVarcharSQL)
			lnResult = sqlexec(lnHandle, lcVarcharSQL)
			if lnResult < 0
			* handle error
			endif
		endif not empty(lcVarcharSQL)
	    sqldisconnect(lnHandle)
	else
	* handle error
	endif
endif
