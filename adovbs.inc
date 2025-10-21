<%
'================================================
' ADO constants include file for VBScript (ASP Classic)
' Valores hexadecimais convertidos para decimais
'================================================

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- CursorOptionEnum Values ----
Const adHoldRecords = 256          ' 0x00000100
Const adMovePrevious = 512         ' 0x00000200
Const adAddNew = 16781312          ' 0x01000400
Const adDelete = 16785408          ' 0x01000800
Const adUpdate = 16797696          ' 0x01008000
Const adBookmark = 8192            ' 0x00002000
Const adApproxPosition = 16384     ' 0x00004000
Const adUpdateBatch = 65536        ' 0x00010000
Const adResync = 131072            ' 0x00020000
Const adNotify = 262144            ' 0x00040000
Const adFind = 524288              ' 0x00080000
Const adSeek = 4194304             ' 0x00400000
Const adIndex = 8388608            ' 0x00800000

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- ExecuteOptionEnum Values ----
Const adAsyncExecute = 16          ' 0x00000010
Const adAsyncFetch = 32            ' 0x00000020
Const adAsyncFetchNonBlocking = 64 ' 0x00000040
Const adExecuteNoRecords = 128     ' 0x00000080
Const adExecuteStream = 1024       ' 0x00000400
Const adExecuteRecord = 2048       ' 0x00000800

'---- DataTypeEnum Values ----
Const adEmpty = 0
Const adTinyInt = 16
Const adSmallInt = 2
Const adInteger = 3
Const adBigInt = 20
Const adUnsignedTinyInt = 17
Const adUnsignedSmallInt = 18
Const adUnsignedInt = 19
Const adUnsignedBigInt = 21
Const adSingle = 4
Const adDouble = 5
Const adCurrency = 6
Const adDecimal = 14
Const adNumeric = 131
Const adBoolean = 11
Const adError = 10
Const adUserDefined = 132
Const adVariant = 12
Const adIDispatch = 9
Const adIUnknown = 13
Const adGUID = 72
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205
Const adChapter = 136
Const adFileTime = 64
Const adPropVariant = 138
Const adVarNumeric = 139

'---- ParameterDirectionEnum Values ----
Const adParamUnknown = 0
Const adParamInput = 1
Const adParamOutput = 2
Const adParamInputOutput = 3
Const adParamReturnValue = 4

'---- CommandTypeEnum Values ----
Const adCmdUnknown = 8             ' &H0008
Const adCmdText = 1                ' &H0001
Const adCmdTable = 2               ' &H0002
Const adCmdStoredProc = 4          ' &H0004
Const adCmdFile = 256              ' &H0100
Const adCmdTableDirect = 512       ' &H0200

'---- ConnectOptionEnum Values ----
Const adAsyncConnect = 16          ' 0x00000010

'---- ObjectStateEnum Values ----
Const adStateClosed = 0
Const adStateOpen = 1
Const adStateConnecting = 2
Const adStateExecuting = 4
Const adStateFetching = 8

'---- RecordStatusEnum Values ----
Const adRecOK = 0
Const adRecNew = 1
Const adRecModified = 2
Const adRecDeleted = 4
Const adRecUnmodified = 8
Const adRecInvalid = 16
Const adRecMultipleChanges = 64
Const adRecPendingChanges = 128
Const adRecCanceled = 256
Const adRecCantRelease = 1024
Const adRecConcurrencyViolation = 2048
Const adRecIntegrityViolation = 4096
Const adRecMaxChangesExceeded = 8192
Const adRecObjectOpen = 16384
Const adRecOutOfMemory = 32768
Const adRecPermissionDenied = 65536
Const adRecSchemaViolation = 131072
Const adRecDBDeleted = 262144
%>