Attribute VB_Name = "utl_functions"
'
' utl_functions  1.1  2/21/2003
'
'  Access Client Solutions Utility Bridge
'                   (C) COPYRIGHT IBM CORP. 2002, 2016
'
' All rights reserved.  Provided on an "AS IS" basis, no
' warranty expressed or implied.
'

' Example of Utility function call.
'  The utl_GetString function will get a string from the host in the row and column specified
'  and return the status of the function call indicating success or type of problem.
'
'  utl_GetString parameters:
'    intRow             the row position of the string in the host session.
'    intColumn          the column position of the string in the host session.
'    strStringBuffer    the buffer for the string; must be defined and allocated by the user.
'    intStringLength    the length of the string.
'
' function definition
'  Public Function utl_GetString(intRow As Integer, intColumn As Integer, strStringBuffer As String, intStringLength As Integer) As Long
' function number call
'  HllFunctionNo = UA_GET_STRING '235
' data buffer
'  HllData = ""
' length of string
'  HllLength = intStringLength
' return code
'  HllReturnCode = 0
' row position
'  HllParm5 = intRow
' column position
'  HllParm6 = intColumn
' function call
'  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
' put the string into the user supplied buffer
'  strStringBuffer = Mid(HllData, 1, intStringLength)
' set the return code
'  utl_GetString = HllReturnCode
' end the function
'  End Function

Option Explicit

Public Function utl_GetWindowHandle(sessionName As String) As Long
  On Error Resume Next
  printTrace ("utl_GetWindowHandle: sessionName= " & sessionName)
  HllFunctionNo = UA_GET_WINDOW_HANDLE '278
  HllData = Trim(sessionName)
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetWindowHandle = HllReturnCode
  printTrace ("...rc= " & utl_GetWindowHandle)
End Function

Public Function utl_GetWindowTitle(sessionName As String) As String
  Dim title As String
  On Error Resume Next
  printTrace ("utl_GetWindowTitle: sessionName= " & sessionName)
    HllFunctionNo = UA_GET_WINDOW_HANDLE '278
  HllData = Trim(sessionName)
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  If HllReturnCode > 0 Then
  Dim txtLen As Long
  txtLen = GetWindowTextLength(HllReturnCode) + 1
  title = Space(txtLen)
    lngTempRC = GetWindowText(HllReturnCode, title, txtLen)
    utl_GetWindowTitle = title
  Else
    utl_GetWindowTitle = ""
  End If
  printTrace ("...rc= " & utl_GetWindowTitle)
End Function

Public Function utl_ConnectSession(sessionName As String) As Long
  Dim strSessionLetter As String
  On Error Resume Next
  
   ' set debug values
  blnDebug = False       ' for msgbox output
  blnWriteToFile = False ' for file output
  
  ' Delete trace files
  If blnDebug Then
    Kill strUtlVBtrace
    Kill "c:\debug.log"
    Kill "c:\utlJavaTrace.txt"
  End If
  
  printTrace ("utl_ConnectSession: sessionName= " & sessionName)
  strSessionLetter = Trim(sessionName)
  HllFunctionNo = UA_CONNECT_SESSION '214
  HllData = strSessionLetter
  HllLength = 4         ' must be 4
  HllReturnCode = 0     ' NA on function call
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  If HllReturnCode = 1 Then  ' OK '
    ConnectSession = strSessionLetter
  End If
  utl_ConnectSession = HllReturnCode
  printTrace ("...rc= " & utl_ConnectSession)
End Function

Public Function utl_DisconnectSession() As Long
  printTrace ("utl_DisconnectSession")
  HllFunctionNo = UA_DISCONNECT_SESSION '216
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  If HllReturnCode = 1 Then  ' OK '
    ConnectSession = ""
  End If
  utl_DisconnectSession = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_RowColumn(intPosition As Integer, intOption As Integer) As Long
  printTrace ("utl_RowColumn: intPosition = " & intPosition & " intOption= " & intOption)
  HllFunctionNo = UA_ROW_COLUMN '247
  HllData = ConnectSession & "000P000"   ' using ehllapi format
  HllLength = intOption
  HllReturnCode = intPosition
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  ' java returns row or column in RC
  utl_RowColumn = HllReturnCode
End Function

Public Function utl_SendKey(strKeys As String) As Long
  printTrace ("utl_SendKey: strKeys= " & strKeys)
  HllFunctionNo = UA_SEND_KEY '255
  HllData = strKeys
  HllLength = Len(strKeys)
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_SendKey = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'*******************************************************'
'*** This method is for compatibility with OLE calls ***'
'*******************************************************'
Public Function utl_SendKeys(strKeys As String) As Long
  printTrace ("utl_SendKeys: strKeys= " & strKeys)
  HllFunctionNo = UA_SEND_KEY '255
  HllData = strKeys
  HllLength = Len(strKeys)
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_SendKeys = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_SendString(intRow As Integer, intColumn As Integer, strSendString As String) As Long
  printScreen
  printTrace ("utl_SendString: intRow= " & intRow & " intColumn= " & intColumn)
  HllFunctionNo = UA_SEND_STRING '256
  HllData = strSendString
  HllLength = Len(strSendString)
  HllReturnCode = 0
  HllParm5 = intRow
  HllParm6 = intColumn
  printTrace5
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_SendString = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_GetString(intRow As Integer, intColumn As Integer, strGetString As String, intStringLength As Integer) As Long
  printTrace ("utl_GetString: intRow= " & intRow & " intColumn= " & intColumn & " intStringLength= " & intStringLength)
  HllFunctionNo = UA_GET_STRING '235
  HllData = ""
  HllLength = intStringLength
  HllReturnCode = 0
  HllParm5 = intRow
  HllParm6 = intColumn
  printTrace5
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  strGetString = Mid(HllData, 1, intStringLength)
  utl_GetString = HllReturnCode
  printTrace ("strGetString= " & strGetString)
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_GetStringFromField(intRow As Integer, intColumn As Integer, strGetString As String, intStringLength As Integer) As Long
  printTrace ("utl_GetStringFromField: intRow= " & intRow & " intColumn= " & intColumn & " intStringLength= " & intStringLength)
  HllFunctionNo = UA_GET_STRING_FROM_FIELD '236
  HllData = ""
  HllLength = intStringLength
  HllReturnCode = 0
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  strGetString = Mid(HllData, 1, intStringLength)
  utl_GetStringFromField = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_SetCursorLocation(intRow As Integer, intColumn As Integer) As Long
  printTrace ("utl_SetCursorLocation: intRow= " & intRow & " intColumn= " & intColumn)
  HllFunctionNo = UA_SET_CURSOR_LOCATION '260
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_SetCursorLocation = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'*******************************************************'
'*** This method is for compatibility with OLE calls ***'
'*******************************************************'
Public Function utl_MoveTo(intRow As Integer, intColumn As Integer) As Long
  printTrace ("utl_MoveTo: intRow= " & intRow & " intColumn= " & intColumn)
  HllFunctionNo = UA_SET_CURSOR_LOCATION '260
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_MoveTo = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_GetFieldInfo(intRow As Integer, intColumn As Integer, intOption As Integer) As Long
  printTrace ("utl_GetFieldInfo: intRow= " & intRow & " intColumn= " & intColumn & " intOption= " & intOption)
  HllFunctionNo = UA_GET_FIELD_INFO '225
  HllData = ""
  HllLength = 0
  HllReturnCode = intOption
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetFieldInfo = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_SendStringToField(intRow As Integer, intColumn As Integer, strSendString As String) As Long
  printTrace ("utl_SendStringToField: intRow= " & intRow & " intColumn= " & intColumn & " strSendString= " & strSendString)
  HllFunctionNo = UA_SEND_STRING_TO_FIELD '257
  HllData = strSendString
  HllLength = Len(strSendString)
  HllReturnCode = 0
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_SendStringToField = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_SearchSession(intRow As Integer, intColumn As Integer, strSearchString As String, intOption As Integer) As Long
  printTrace ("utl_SearchSession: intRow= " & intRow & " intColumn= " & intColumn & " strSearchString= " & strSearchString & " intOption= " & intOption)
  HllFunctionNo = UA_SEARCH_SESSION '252
  HllData = strSearchString
  HllLength = Len(strSearchString)
  HllReturnCode = intOption
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_SearchSession = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_SearchField(intRow As Integer, intColumn As Integer, strSearchString As String) As Long
  printTrace ("utl_SearchField: intRow= " & intRow & " intColumn= " & intColumn & " strSearchString= " & strSearchString)
  HllFunctionNo = UA_SEARCH_FIELD '252
  HllData = strSearchString
  HllLength = Len(strSearchString)
  HllReturnCode = 0
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_SearchField = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' Only options gbl_XSTATUS, gbl_CONNECTION and gbl_ERROR are supported this function.
'
Public Function utl_GetConnectionStatus(intStatusType As Integer) As Long
  printTrace ("utl_GetConnectionStatus: intStatusType= " & intStatusType)
  HllFunctionNo = UA_GET_CONNECTION_STATUS '220
  HllData = ""
  HllLength = 0
  HllReturnCode = intStatusType
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetConnectionStatus = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_Wait() As Long
  printTrace ("utl_Wait")
  HllFunctionNo = UA_WAIT  '269
  HllData = 0
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_Wait = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_WaitHostQuiet(intSettleTime As Integer, intTimeout As Integer) As Long
  printTrace ("utl_WaitHostQuiet: intSettleTime= " & intSettleTime & " intTimeOut= " & intTimeout)
  HllFunctionNo = UA_WAIT_HOST_QUIET '277
  HllData = ""
  HllLength = intSettleTime
  HllReturnCode = intTimeout
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_WaitHostQuiet = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_WaitForString(intRow As Integer, intColumn As Integer, strWaitString As String, intTimeout As Integer) As Long
  printTrace ("utl_WaitForString: intRow= " & intRow & " intColumn= " & intColumn & " strWaitString= " & strWaitString & " intTimeout= " & intTimeout)
  HllFunctionNo = UA_WAIT_FOR_STRING '276
  HllData = strWaitString
  HllLength = Len(strWaitString)
  HllReturnCode = intTimeout
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_WaitForString = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_GetSessionSize() As Long
  printTrace ("utl_GetSessionSize")
  HllFunctionNo = UA_GET_SESSION_SIZE  '233
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetSessionSize = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' This function returns a 1 for all options if the session is connected.
'
Public Function utl_GetSessions(strSessionList As String, intListLength As Integer, intSessionState As Integer) As Long
  printTrace ("utl_GetSessions: strSessionList= " & strSessionList & " intListLength= " & intListLength & " intSessionState= " & intSessionState)
  HllFunctionNo = UA_GET_SESSIONS '232
  HllData = strSessionList
  HllLength = intListLength
  HllReturnCode = intSessionState
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetSessions = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_RegisterClient(intType As Integer) As Long
  printTrace ("utl_RegisterClient")
  HllFunctionNo = UA_REGISTER_CLIENT '244
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_RegisterClient = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_UnRegisterClient() As Long
  printTrace ("utl_UnRegisterClient")
  HllFunctionNo = UA_UNREGISTER_CLIENT '268
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_UnRegisterClient = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_ResetSystem() As Long
  printTrace ("utl_ResetSystem")
  HllFunctionNo = UA_RESET_SYSTEM '245
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_ResetSystem = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_Pause(intHalfSec As Integer) As Long
  printTrace ("utl_Pause")
  HllFunctionNo = UA_PAUSE '242
  HllData = ""
  HllLength = intHalfSec
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_Pause = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'  utl_StartSession takes two parameters
'    the url of the HOD session (Example: http://server/HOD/utl.html)
'    a 1-char option to show how to start the session
'      Normal, Icon, Hidden, Maximized
'  The utlBridge call passes one string which consists of the two parms separated by a comma.
'
Public Function utl_StartSession(strURL As String, strOption As String) As Long
  Dim strURLOption As String
  strURLOption = strURL & "," & strOption
   
  printTrace ("utl_StartSession")
  HllFunctionNo = UA_START_SESSION '264
  HllData = strURLOption
  HllLength = Len(strURLOption)
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_StartSession = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_StartKeystrokeIntercept(strSession As String, intFilter As Integer) As Long
  printTrace ("utl_StartKeystrokeIntercept: strSession= " & strSession & " intFilter= " & intFilter)
  HllFunctionNo = UA_START_KEYSTROKE_INTERCEPT  '263
  HllData = strSession
  HllLength = 4
  HllReturnCode = intFilter
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_StartKeystrokeIntercept = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_StopKeystrokeIntercept(strSession As String) As Long
  printTrace ("utl_StopKeystrokeIntercept: strSession= " & strSession)
  HllFunctionNo = UA_STOP_KEYSTROKE_INTERCEPT  '265
  HllData = strSession
  HllLength = 4
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_StopKeystrokeIntercept = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_WaitForHostConnect(intTimeout As Integer) As Long
  printTrace ("utl_WaitForHostConnect:  intTimeout= " & intTimeout)
  HllFunctionNo = UA_WAIT_FOR_HOST_CONNECT  '273
  HllData = ""
  HllLength = 0
  HllReturnCode = intTimeout
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_WaitForHostConnect = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_SetParameter(intParm As Integer, intSetting As Integer, strEscapeChar As String) As Long
  ' when setting trace, 1 is on, 2 is off
  printTrace ("utl_SetParameter")
  HllFunctionNo = UA_SET_PARAMETER '261
  HllData = strEscapeChar
  HllLength = intParm
  HllReturnCode = intSetting
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_SetParameter = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_GetEmulatorVersion(strResult As String, intLen As Integer) As Long
  printTrace ("utl_GetEmulatorVersion")
  HllFunctionNo = UA_GET_EMULATOR_VERSION '223
  HllData = ""
  HllLength = intLen
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetEmulatorVersion = HllReturnCode
  strResult = Trim(HllData)
  printTrace ("...rc= " & HllReturnCode & " version= " & strResult)
End Function

Public Function utl_GetCursorLocation() As Long
  printTrace ("utl_GetCursorLocation")
  HllFunctionNo = UA_GET_CURSOR_LOCATION '221
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetCursorLocation = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_GetFieldLength(intRow As Integer, intColumn As Integer, intType As Integer) As Long
  printTrace ("utl_GetFieldLength: intRow= " & intRow & " intColumn= " & intColumn & " intType= " & intType)
  HllFunctionNo = UA_GET_FIELD_LENGTH '226
  HllData = ""
  HllLength = 0
  HllReturnCode = intType
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetFieldLength = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_GetFieldPosition(intRow As Integer, intColumn As Integer, intType As Integer) As Long
  printTrace ("utl_GetFieldPosition: intRow= " & intRow & " intColumn= " & intColumn & " intType= " & intType)
  HllFunctionNo = UA_GET_FIELD_POSITION '227
  HllData = ""
  HllLength = 0
  HllReturnCode = intType
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetFieldPosition = HllReturnCode
  printTrace ("...rc= " & HllReturnCode & " pos= " & HllLength)
  
End Function

Public Function utl_GetParameter(intParmIndex As Integer) As Long
  printTrace ("utl_GetParameter: intParmIndex= " & intParmIndex)
  HllFunctionNo = UA_GET_PARAMETER  '230
  HllData = ""
  HllLength = 0
  HllReturnCode = intParmIndex
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetParameter = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_SetTrace(intTrace As Integer) As Long
' Trace on = 1, off = 2
printTrace ("utl_SetTrace: intTrace= " & intTrace)
  HllFunctionNo = UA_SET_TRACE '290
  HllData = ""
  HllLength = 0
  HllReturnCode = intTrace
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_SetTrace = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' This function returns a 1 for all options if the session is connected.
'
Public Function utl_ListSessions(strBuffer As String, intBufferLength As Integer, strTitle As String, intType As Integer) As Long
  printTrace ("utl_ListSessions")
  HllFunctionNo = UA_LIST_SESSIONS  '238
  HllData = strBuffer
  HllLength = intBufferLength
  HllReturnCode = intType
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_ListSessions = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_LockKeyboard() As Long
  printTrace ("utl_LockKeyboard")
  HllFunctionNo = UA_LOCK_KEYBOARD  '239
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_LockKeyboard = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_UnlockKeyboard() As Long
  printTrace ("utl_UnlockKeyboard")
  HllFunctionNo = UA_UNLOCK_KEYBOARD  '267
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_UnlockKeyboard = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_GetEmulatorPath(strBuffer As String, intBufferLength As Integer) As Long
  printTrace ("utl_GetEmulatorPath")
  'HllFunctionNo = UA_GET_EMULATOR_PATH  '222
  HllReturnCode = 1
  utl_GetEmulatorPath = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_AllowUpdates() As Long
  printTrace ("utl_AllowUpdates")
  'HllFunctionNo = UA_ALLOW_UPDATES  '210
  HllReturnCode = 1
  utl_AllowUpdates = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_BlockUpdates() As Long
  printTrace ("utl_BlockUpdates")
  'HllFunctionNo = UA_BLOCK_UPDATES  '211
  HllReturnCode = 1
  utl_BlockUpdates = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_OpenConfiguration(strPathConfigFileName As String) As Long
  printTrace ("utl_OpenConfiguration")
  'HllFunctionNo = UA_OPEN_CONFIGURATION  '240
  HllReturnCode = 1
  utl_OpenConfiguration = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_GetConfiguration(strBuffer As String, intBufferLength As Integer) As Long
  printTrace ("utl_GetConfiguration")
  utl_GetConfiguration = 1
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_CloseConfiguration() As Long
  printTrace ("utl_CloseConfiguration")
  'HllFunctionNo = UA_CLOSE_CONFIGURATION  '213
  HllReturnCode = 1
  utl_CloseConfiguration = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_GetSessionHandle(strHandle As String) As Long
  printTrace ("utl_GetSessionHandle")
  'HllFunctionNo = UA_GET_SESSION_HANDLE  '231
  HllReturnCode = 1
  utl_GetSessionHandle = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_GetLayoutName(strBuffer As String, intBufferLength As Integer) As Long
  printTrace ("utl_GetLayoutName")
  'HllFunctionNo = UA_GET_LAYOUT_NAME  '229
  HllReturnCode = 1
  utl_GetLayoutName = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_OpenLayout(strPathLayoutFile As String) As Long
  printTrace ("utl_OpenLayout")
  'HllFunctionNo = UA_OPEN_LAYOUT  '241
  utl_OpenLayout = 1
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' This function returns a 1 for all options if the session is ready.
'
Public Function utl_GetSessionStatus(strSession As String, intType As Integer) As Long
  printTrace ("utl_GetSessionStatus: strSession= " & strSession & " intType= " & intType)
  HllFunctionNo = UA_GET_SESSION_STATUS  '234
  HllData = strSession
  HllLength = 4
  HllReturnCode = intType
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_GetSessionStatus = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_GetError(strBuffer As String, intBufferLength As Integer) As Long
 printTrace ("utl_GetError")
 HllReturnCode = 1
 printTrace ("...rc= " & HllReturnCode)
 utl_GetError = HllReturnCode
End Function

'
' Runs the previously created HOD macro on the specified session.
'
 Public Function utl_RunEmulatorMacro(strSession As String, strMacroName As String) As Long
   printTrace ("utl_RunEmulatorMacro")
   HllFunctionNo = UA_RUN_EMULATOR_MACRO  '248
   HllData = strSession & strMacroName
   HllLength = Len(strSession & strMacroName)
   HllReturnCode = 0
   HllParm5 = 0
   HllParm6 = 0
   printTraceAll
   lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
   printTrace ("...rc= " & HllReturnCode)
   utl_RunEmulatorMacro = HllReturnCode
 End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_SessionOff() As Long
  printTrace ("utl_SessionOff")
  HllFunctionNo = UA_SESSION_OFF  '258
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  printTrace ("...rc= " & HllReturnCode)
  utl_SessionOff = HllReturnCode
End Function
    
'
' For compatibility, this function returns a 1.
'
Public Function utl_SessionOn() As Long
  printTrace ("utl_SessionOn")
  HllFunctionNo = UA_SESSION_ON  '259
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  printTrace ("...rc= " & HllReturnCode)
  utl_SessionOn = HllReturnCode
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_ShowLastError() As Long
 printTrace ("utl_ShowLastError")
 HllReturnCode = 1
 printTrace ("...rc= " & HllReturnCode)
 utl_ShowLastError = HllReturnCode
End Function

Public Function utl_WaitForCursor(intRow As Integer, intColumn As Integer, intTimeout As Integer) As Long
  printTrace ("utl_WaitForCursor: intRow= " & intRow & " intColumn= " & intColumn & " intTimeout= " & intTimeout)
  HllFunctionNo = UA_WAIT_FOR_CURSOR  '270
  HllData = ""
  HllLength = 0
  HllReturnCode = intTimeout
  HllParm5 = intRow
  HllParm6 = intColumn
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  printTrace ("...rc= " & HllReturnCode)
  utl_WaitForCursor = HllReturnCode
End Function

Public Function utl_WaitForCursorMove(intTimeout As Integer) As Long
  printTrace ("utl_WaitForCursorMove: intTimeout= " & intTimeout)
  HllFunctionNo = UA_WAIT_FOR_CURSOR_MOVE  '271
  HllData = ""
  HllLength = 0
  HllReturnCode = intTimeout
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  printTrace ("...rc= " & HllReturnCode)
  utl_WaitForCursorMove = HllReturnCode
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_WaitForHostDisconnect(intTimeout As Integer) As Long
  printTrace ("utl_WaitForHostDisconnect: intTimeout= " & intTimeout)
  HllFunctionNo = UA_WAIT_FOR_HOST_DISCONNECT  '274
  HllData = ""
  HllLength = 0
  HllReturnCode = intTimeout
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  printTrace ("...rc= " & HllReturnCode)
  utl_WaitForHostDisconnect = HllReturnCode
End Function

Public Function utl_WaitForKey(strKey As String, intTimeout As Integer) As Long
  printTrace ("utl_WaitForKey: strKey= " & strKey & " intTimeout= " & intTimeout)
  HllFunctionNo = UA_WAIT_FOR_KEY  '275
  HllData = strKey
  HllLength = Len(strKey)
  HllReturnCode = intTimeout
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  printTrace ("...rc= " & HllReturnCode)
  utl_WaitForKey = HllReturnCode
End Function

Public Function utl_SendAndWait(intRow As Integer, intColumn As Integer, strSendString As String, strWaitString As String, intTimeout As Integer) As Long
  Dim intWaitStringLength As Integer
  printTrace ("utl_SendAndWait: intRow= " & intRow & " intColumn= " & intColumn & " strSendString= " & strSendString & " strWaitString= " & strWaitString & " intTimeout= " & intTimeout)
  HllFunctionNo = UA_SEND_AND_WAIT  '253
  HllData = strSendString & "," & strWaitString
  HllLength = Len(strSendString & "," & strWaitString)
  HllReturnCode = intTimeout
  HllParm5 = intRow
  HllParm6 = intColumn
  printTrace5
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_SendAndWait = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_ReceiveFile(strBuffer As String, intBufferLength As Integer) As Long
  printTrace ("utl_ReceiveFile: strBuffer= " & strBuffer & " intBufferLength= " & intBufferLength)
  HllFunctionNo = UA_RECEIVE_FILE  '243
  HllData = strBuffer
  HllLength = Len(strBuffer)
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  printTrace ("...rc= " & HllReturnCode)
  utl_ReceiveFile = HllReturnCode
End Function

Public Function utl_SendFile(strBuffer As String, intBufferLength As Integer) As Long
  ' test for vm
  'strBuffer = "c:\\temp1.txt  temp1 text a  ( ASCII CRLF RECFM V LRECL 133"
  ' test for mvs
  'strBuffer = "c:\\temp2.txt  user.temp2.text  ASCII CRLF RECFM (V) LRECL (133)"
  
  printTrace ("utl_SendFile: strBuffer= " & strBuffer & " intBufferLength= " & intBufferLength)
  HllFunctionNo = UA_SEND_FILE  '254
  HllData = strBuffer
  HllLength = Len(strBuffer)
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  printTrace ("...rc= " & HllReturnCode)
  utl_SendFile = HllReturnCode
End Function

Public Function utl_StopSession() As Long
  printTrace ("utl_StopSession")
  HllFunctionNo = UA_STOP_SESSION  '266
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  printTrace ("...rc= " & HllReturnCode)
  utl_StopSession = HllReturnCode
End Function

Public Function utl_AddWait(intTable As Integer, intEvent As Integer) As Long
  printTrace ("utl_AddWait: intTable= " & intTable & " intEvent= " & intEvent)
  HllFunctionNo = UA_ADD_WAIT  '201
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = intTable
  HllParm6 = intEvent
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_AddWait = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_AddWaitForCursor(intTable As Integer, intEvent As Integer, intRow As Integer, intColumn As Integer) As Long
  Dim rrrrcccc As Long
  printTrace ("utl_AddWaitForCursor: intTable= " & intTable & " intEvent= " & intEvent & " intRow= " & intRow & " intColumn= " & intColumn)
  HllFunctionNo = UA_ADD_WAIT_FOR_CURSOR  '202
  HllData = ""
  HllLength = 0
  rrrrcccc = (intRow * 2 ^ 16) + intColumn
  printTrace ("hex(rrrrcccc)= " & Hex(rrrrcccc))
  HllReturnCode = rrrrcccc
    'Row in high order bytes
    'Column in low order bytes
  HllParm5 = intTable
  HllParm6 = intEvent
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_AddWaitForCursor = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_AddWaitForCursorMove(intTable As Integer, intEvent As Integer) As Long
  printTrace ("utl_AddWaitForCursorMove: intTable= " & intTable & " intEvent= " & intEvent)
  HllFunctionNo = UA_ADD_WAIT_FOR_CURSOR_MOVE  '203
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = intTable
  HllParm6 = intEvent
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_AddWaitForCursorMove = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_AddWaitForHostConnect(intTable As Integer, intEvent As Integer) As Long
  printTrace ("utl_AddWaitForHostConnect: intTable= " & intTable & " intEvent= " & intEvent)
  HllFunctionNo = UA_ADD_WAIT_FOR_HOST_CONNECT  '204
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = intTable
  HllParm6 = intEvent
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_AddWaitForHostConnect = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

'
' For compatibility, this function returns a 1.
'
Public Function utl_AddWaitForHostDisconnect(intTable As Integer, intEvent As Integer) As Long
  printTrace ("utl_AddWaitForHostDisconnect: intTable= " & intTable & " intEvent= " & intEvent)
  HllFunctionNo = UA_ADD_WAIT_FOR_HOST_DISCONNECT  '205
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = intTable
  HllParm6 = intEvent
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_AddWaitForHostDisconnect = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_AddWaitForKey(intTable As Integer, intEvent As Integer, strKey As String) As Long
  printTrace ("utl_AddWaitForKey: intTable= " & intTable & " intEvent= " & intEvent & " strKey= " & strKey)
  HllFunctionNo = UA_ADD_WAIT_FOR_KEY  '206
  HllData = strKey
  HllLength = Len(strKey)
  HllReturnCode = 0
  HllParm5 = intTable
  HllParm6 = intEvent
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_AddWaitForKey = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_AddWaitForString(intTable As Integer, intEvent As Integer, intRow As Integer, intColumn As Integer, strString As String) As Long
  Dim rrrrcccc As Long
  printTrace ("utl_AddWaitForString: intTable= " & intTable & " intEvent= " & intEvent & " intRow= " & intRow & " intColumn= " & intColumn & " strString= " & strString)
  HllFunctionNo = UA_ADD_WAIT_FOR_STRING  '207
  HllData = strString
  HllLength = Len(strString)
  rrrrcccc = (intRow * 2 ^ 16) + intColumn
  printTrace ("hex(rrrrcccc)= " & Hex(rrrrcccc))
  HllReturnCode = rrrrcccc
  HllParm5 = intTable
  HllParm6 = intEvent
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_AddWaitForString = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_AddWaitForStringNotAt(intTable As Integer, intEvent As Integer, intRow As Integer, intColumn As Integer, strString As String) As Long
  Dim rrrrcccc As Long
  printTrace ("utl_AddWaitForStringNotAt: intTable= " & intTable & " intEvent= " & intEvent & " intRow= " & intRow & " intColumn= " & intColumn & " strString= " & strString)
  HllFunctionNo = UA_ADD_WAIT_FOR_STRING_NOT_AT  '208
  HllData = strString
  HllLength = Len(strString)
  rrrrcccc = (intRow * 2 ^ 16) + intColumn
  printTrace ("hex(rrrrcccc)= " & Hex(rrrrcccc))
  HllReturnCode = rrrrcccc
  HllParm5 = intTable
  HllParm6 = intEvent
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_AddWaitForStringNotAt = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_AddWaitHostQuiet(intTable As Integer, intEvent As Integer, intSettleTime As Integer) As Long
  printTrace ("utl_AddWaitHostQuiet: intTable= " & intTable & " intEvent= " & intEvent & " intSettleTime= " & intSettleTime)
  HllFunctionNo = UA_ADD_WAIT_HOST_QUIET  '209
  HllData = ""
  HllLength = 0
  HllReturnCode = intSettleTime
  HllParm5 = intTable
  HllParm6 = intEvent
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_AddWaitHostQuiet = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_ClearEventTable(intTable As Integer) As Long
  printTrace ("utl_ClearEventTable: intTable= " & intTable)
  HllFunctionNo = UA_CLEAR_EVENT_TABLE  '212
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = intTable
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_ClearEventTable = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_DeleteEvent(intTable As Integer, intEvent As Integer) As Long
  printTrace ("utl_DeleteEvent: intTable= " & intTable & " intEvent= " & intEvent)
  HllFunctionNo = UA_DELETE_EVENT  '215
  HllData = ""
  HllLength = 0
  HllReturnCode = 0
  HllParm5 = intTable
  HllParm6 = intEvent
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_DeleteEvent = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_WaitForEvent(intTable As Integer, intTimeout As Integer) As Long
  printTrace ("utl_WaitForEvent: intTable= " & intTable & " intTimeout= " & intTimeout)
  HllFunctionNo = UA_WAIT_FOR_EVENT  '275
  HllData = ""
  HllLength = 0
  HllReturnCode = intTimeout
  HllParm5 = intTable
  HllParm6 = 0
  printTraceAll
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  utl_WaitForEvent = HllReturnCode
  printTrace ("...rc= " & HllReturnCode)
End Function

Public Function utl_GetKeyStroke(strSession As String, strKeyBuffer As String, intKeyBufferLength As Integer) As Long
  printTrace ("utl_GetKeyStroke: strSession= " & strSession & " strKeyBuffer= " & strKeyBuffer & " intKeyBufferLength= " & intKeyBufferLength)
  HllFunctionNo = UA_GET_KEY_STROKE '228
  ' pass session in buffer area
  strKeyBuffer = strSession & strKeyBuffer
  HllData = strKeyBuffer
  ' check for minimum buffer length
  If intKeyBufferLength < 12 Then intKeyBufferLength = 12
  HllLength = intKeyBufferLength
  HllReturnCode = 0
  HllParm5 = 0
  HllParm6 = 0
  printTrace5
  lngTempRC = utlapi&(HllFunctionNo, HllData, HllLength, HllReturnCode, HllParm5, HllParm6)
  strKeyBuffer = Trim(HllData)
  utl_GetKeyStroke = HllReturnCode
  printTrace ("utl_GetKeyStroke= " & strKeyBuffer)
  printTrace ("...rc= " & HllReturnCode)
End Function

