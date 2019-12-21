Attribute VB_Name = "UTLAPI_Declares"
'
' IBMGLOBALS  1.0  07/18/2002
'
' Access Client Solutions Utility Bridge
'                   (C) COPYRIGHT IBM CORP. 2002, 2016
'
' All rights reserved.  Provided on an "AS IS" basis, no
' warranty expressed or implied.
'
Declare Function UTLAddWait Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nEventID%) As Long
Declare Function UTLAddWaitForCursor Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nEventID%, ByVal nRow%, ByVal nColumn%) As Long
Declare Function UTLAddWaitForCursorMove Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nEventID%) As Long
Declare Function UTLAddWaitForHostConnect Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nEventID%) As Long
Declare Function UTLAddWaitForHostDisconnect Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nEventID%) As Long
Declare Function UTLAddWaitForKey Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nEventID%, ByVal lpKey$) As Long
Declare Function UTLAddWaitForString Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nEventID%, ByVal nRow%, ByVal nColumn%, ByVal lpString$) As Long
Declare Function UTLAddWaitForStringNotAt Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nEventID%, ByVal nRow%, ByVal nColumn%, ByVal lpString$) As Long
Declare Function UTLAddWaitHostQuiet Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nEventID%, ByVal nIdleMilliseconds%) As Long
Declare Function UTLAllowUpdates Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLBlockUpdates Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLClearEventTable Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%) As Long
Declare Function UTLCloseConfiguration Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLConnectSession Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Session$) As Long
Declare Function UTLDeleteEvent Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nEventID%) As Long
Declare Function UTLDisconnectSession Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLExecute Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal MacroCode$, ByVal nTimeout%) As Long
Declare Function UTLFileTransferDlg Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal A1$, ByVal A2$, ByVal A3$, ByVal A4$, ByVal Num1%, ByVal A5$, ByVal Num2%) As Long
Declare Function UTLGet_____Path Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Path$, ByVal Length%) As Long
Declare Function UTLGet_____Version Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal buffer$, ByVal Length%) As Long
Declare Function UTLGetConfiguration Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal config$, ByVal Length%) As Long
Declare Function UTLGetConnectionStatus Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal StatusType%) As Long
Declare Function UTLGetCursorLocation Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLGetEmulatorPath Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Path$, ByVal Length%) As Long
Declare Function UTLGetEmulatorVersion Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal buffer$, ByVal Length%) As Long
Declare Function UTLGetError Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal errstring$, ByVal Length%) As Long
Declare Function UTLGetFieldInfo Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Row%, ByVal Column%, ByVal InfoType%) As Long
Declare Function UTLGetFieldLength Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Row%, ByVal Column%, ByVal field%) As Long
Declare Function UTLGetFieldPosition Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Row%, ByVal Column%, ByVal field%) As Long
Declare Function UTLGetKeyStroke Lib "ATMAPI32.DLL"  (ByVal hWnd&, ByVal Session$, ByVal GetStroke$, ByVal BufferLength%) As Long
Declare Function UTLGetLayoutName Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal layout$, ByVal Length%) As Long
Declare Function UTLGetParameter Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Index%) As Long
Declare Function UTLGetSessionHandle Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Session$) As Long
Declare Function UTLGetSessions Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal SessionList$, ByVal Length%, ByVal state%) As Long
Declare Function UTLGetSessionSize Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLGetSessionStatus Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Session$, ByVal status%) As Long
Declare Function UTLGetString Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Row%, ByVal Column%, ByVal GetString$, ByVal Length%) As Long
Declare Function UTLGetStringFromField Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Row%, ByVal Column%, ByVal GetString$, ByVal Length%) As Long
Declare Function UTLGetUTLAPIVersion Lib "ATMAPI32.DLL" (ByVal buffer$, ByVal Length%) As Long
Declare Function UTLHoldHost Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLListSessions Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Session$, ByVal Length%, ByVal title$, ByVal status%) As Long
Declare Function UTLLockKeyboard Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLOpenConfiguration Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal config$) As Long
Declare Function UTLOpenLayout Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal layout$) As Long
Declare Function UTLPasswordDlg Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Password$, ByVal Length%) As Long
Declare Function UTLPause Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Pause%) As Long
Declare Function UTLReceiveFile Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal receive$, ByVal Length%) As Long
Declare Function UTLRegisterClient Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nType%) As Long
Declare Function UTLResetSystem Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLResumeHost Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLRowColumn Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Position%, ByVal Toggle%) As Long
Declare Function UTLRun_____Macro Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Session$, ByVal Macro$) As Long
Declare Function UTLRun_____MacroAsync Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Session$, ByVal Macro$) As Long
Declare Function UTLRunEmulatorMacro Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Session$, ByVal Macro$) As Long
Declare Function UTLSearchField Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Row%, ByVal Column%, ByVal SearchString$) As Long
Declare Function UTLSearchSession Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Row%, ByVal Column%, ByVal SearchString$, ByVal SearchOption%) As Long
Declare Function UTLSendAndWait Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nRow%, ByVal nColumn%, ByVal SendString$, ByVal WaitString$, ByVal nTimeout%) As Long
Declare Function UTLSendFile Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Send$, ByVal Length%) As Long
Declare Function UTLSendKey Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Key$) As Long
Declare Function UTLSendString Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Row%, ByVal Column%, ByVal SendString$) As Long
Declare Function UTLSendStringToField Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Row%, ByVal Column%, ByVal SendString$) As Long
Declare Function UTLSessionOff Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLSessionOn Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLSetCursorLocation Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Row%, ByVal Column%) As Long
Declare Function UTLSetParameter Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Index%, ByVal Setting%, ByVal Escape$) As Long
Declare Function UTLShowLastError Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLStartKeyStrokeIntercept Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Session$, ByVal Intercept%) As Long
Declare Function UTLStartSession Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Session$, ByVal WindowMode$) As Long
Declare Function UTLStopKeyStrokeIntercept Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Session$) As Long
Declare Function UTLStopSession Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLTerminalType Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal ShortName$, ByVal buffer$, ByVal BufferLength%) As Long
Declare Function UTLUnlockKeyboard Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLUnregisterClient Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLWait Lib "ATMAPI32.DLL" (ByVal hWnd&) As Long
Declare Function UTLWaitForCursor Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nRow%, ByVal nColumn%, ByVal nTimeout%) As Long
Declare Function UTLWaitForCursorMove Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTimeout%) As Long
Declare Function UTLWaitForHostConnect Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Timeout%) As Long
Declare Function UTLWaitForHostDisconnect Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Timeout%) As Long
Declare Function UTLWaitForKey Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal Key$, ByVal nTimeout%) As Long
Declare Function UTLWaitForString Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nRow%, ByVal nColumn%, ByVal SearchString$, ByVal nTimeout%) As Long
Declare Function UTLWaitHostQuiet Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal SettleTime%, ByVal Timeout%) As Long
Declare Function UTLWaitForEvent Lib "ATMAPI32.DLL" (ByVal hWnd&, ByVal nTable%, ByVal nTimeout%) As Long

