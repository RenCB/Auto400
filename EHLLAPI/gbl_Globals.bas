Attribute VB_Name = "Module4"
'
' gbl_Globals  1.0  07/18/2002
'
'  Access Client Solutions Utility Bridge
'                   (C) COPYRIGHT IBM CORP. 2002, 2016
'
' All rights reserved.  Provided on an "AS IS" basis, no
' warranty expressed or implied.
'
Option Explicit
'
' Global Constants for gbl_GetSessionStatus
Global Const gbl_ISCONFIGURED = 1
Global Const gbl_ISOPENED = 2
Global Const gbl_ISPOWERED = 3
Global Const gbl_ISEMULATED = 14
Global Const gbl_ISCONNECTED = 15
Global Const gbl_ISFILETRANSFER = 16

' Global Constants for gbl_RowColumn
Global Const gbl_GETROW = 0
Global Const gbl_GETCOLUMN = 1

' Global Constants for gbl_SearchSession, gbl_GetFieldLength and Postion option
Global Const gbl_THISFIELD = 0
Global Const gbl_NEXTFIELD = 1
Global Const gbl_PREVIOUSFIELD = 2
Global Const gbl_NEXTPROTECTEDFIELD = 3
Global Const gbl_NEXTUNPROTECTEDFIELD = 4
Global Const gbl_PREVIOUSPROTECTEDFIELD = 5
Global Const gbl_PREVIOUSUNPROTECTEDFIELD = 6

' Global Constants for gbl_SearchSession option
Global Const gbl_SEARCHALL = 1
Global Const gbl_SEARCHFROM = 2
Global Const gbl_SEARCHAT = 3
Global Const gbl_SEARCHBACK = 4

' Global Constants for Get/SetParameters
Global Const gbl_ATTRIB = 1
Global Const gbl_AUTORESET = 2
Global Const gbl_CONNECTTYPE = 3
Global Const gbl_EAB = 4
Global Const gbl_PAUSE = 5
Global Const gbl_SEARCHORG = 6
Global Const gbl_SEARCHDIRECTION = 7
Global Const gbl_TIMEOUT = 8
Global Const gbl_TRACE = 9
Global Const gbl_WAIT = 10
Global Const gbl_XLATE = 11
Global Const gbl_ESCAPE = 12

' Global Constants for gbl_GetFieldInfo
Global Const gbl_ISFIELDPROTECTED = 1
Global Const gbl_ISFIELDNUMERIC = 2
Global Const gbl_ISFIELDSELECTORPENDETECTABLE = 3
Global Const gbl_ISFIELDBOLD = 4
Global Const gbl_ISFIELDHIDDEN = 5
Global Const gbl_ISFIELDMODIFIED = 6

' Global Constants for gbl_GetSessions and gbl_ListSessions
Global Const gbl_GETCONFIGURED = 1
Global Const gbl_GETOPENED = 2
Global Const gbl_GETPOWERED = 3
Global Const gbl_GETEMULATED = 11
Global Const gbl_GETEMULATEDPOWERED = 12 'Emulated and Powered

' Global Constants for gbl_GetSessions
Global Const gbl_GETCONFIGUREDCOUNT = 4
Global Const gbl_GETOPENEDCOUNT = 5
Global Const gbl_GETPOWEREDCOUNT = 6
Global Const gbl_GETEMULATEDCOUNT = 14
Global Const gbl_GETEMULATEDPOWEREDCOUNT = 15 'Emulated and Powered

' Return codes
Global Const gbl_SUCCESS = 1
Global Const gbl_NOTFOUND = 0
Global Const gbl_NOTATTRIBUTE = 0
Global Const gbl_NOTCONNECTED = -1
Global Const gbl_INVALIDPARAMETER = -2
Global Const gbl_TIMEDOUT = -4
Global Const gbl_SESSIONOCCUPIED = -4
Global Const gbl_SESSIONLOCKED = -5
Global Const gbl_PROTECTED = -5
Global Const gbl_FIELDSIZEMISMATCH = -6
Global Const gbl_DATATRUNCATED = -6
Global Const gbl_INVALIDPOSITION = -7
Global Const gbl_NOPRIORSTARTKEYSTROKE = -8
Global Const gbl_NOPRIORSTARTHOSTNOTIFY = -8
Global Const gbl_SYSTEMERROR = -9
Global Const gbl_FUNCTIONNOTAVAILABLE = -10
Global Const gbl_RESOURCEUNAVAILABLE = -11
Global Const gbl_SESSIONBUSY = -12
Global Const gbl_SEARCHSTRINGNOTFOUND = -24
Global Const gbl_UNFORMATTEDHOSTPS = -24
Global Const gbl_NOSUCHFIELD = -24
Global Const gbl_NOHOSTSESSIONUPDATE = -24
Global Const gbl_KEYSTROKESNOTAVAILABLE = -25
Global Const gbl_HOSTSESSIONUPDATE = -26
Global Const gbl_KEYSTROKEQUEUEOVERFLOW = -31
Global Const gbl_MEMORYUNAVAILABLE = -101
Global Const gbl_DELAYENDEDBYCLIENT = -102
Global Const gbl_UNCONFIGUREDPSID = -103
Global Const gbl_NOEMULATORATTACHED = -104
Global Const gbl_WSCTRLFAILURE = -105
Global Const gbl_NOMATCHINGPSID = -200
Global Const gbl_SESSIONOPEN = -201
Global Const gbl_CONFIGOPEN = -202
Global Const gbl_LIBLOADERROR = -203
Global Const gbl_EVENTALREADYSET = -301
Global Const gbl_EVENTMAXEXCEEDED = -302
Global Const gbl_TABLEMAXEXCEEDED = -303
Global Const gbl_TABLENOTSET = -304
Global Const gbl_INDEXNOTSET = -305
Global Const gbl_INVALIDROW = -306
Global Const gbl_INVALIDCOLUMN = -307
Global Const gbl_STRINGTOOLONG = -308

' Global Constants for gbl_GetConnectionStatus
Global Const gbl_XSTATUS = 1
Global Const gbl_CONNECTION = 2
Global Const gbl_ERROR = 3
Global Const gbl_CASEMODE = 4 'obsolete constant
Global Const gbl_TERMINAL_MODEL = 4
Global Const gbl_CONNECTION_STATUS = 5
Global Const gbl_TRANSMIT_MODE = 6
Global Const gbl_KEYBOARD_LOCK = 7
Global Const gbl_FORMS = 8
Global Const gbl_XMT = 9
Global Const gbl_RCV = 10
Global Const gbl_LTAI = 11

' Return codes for gbl_XSTATUS in gbl_GetConnectionStatus
Global Const gbl_INVALIDNUM = 1
Global Const gbl_NUMONLY = 2
Global Const gbl_PROTFIELD = 3
Global Const gbl_PASTEOF = 4
Global Const gbl_BUSY = 5
Global Const gbl_INVALIDFUNC = 6
Global Const gbl_UNAUTHORIZED = 7
Global Const gbl_SYSTEM = 8
Global Const gbl_INVALIDCHAR = 9

' Return codes for gbl_CONNECTION in gbl_GetConnectionStatus
Global Const gbl_APPOWNED = 1
Global Const gbl_SSCP = 2
Global Const gbl_UNOWNED = 3
Global Const gbl_NONE = 4
Global Const gbl_UNKNOWN = 5

' Return codes for gbl_ERROR in gbl_GetConnectionStatus
Global Const gbl_PROGCHECK = 1
Global Const gbl_COMMCHECK = 2
Global Const gbl_MACHINECHECK = 3

' Return codes for gbl_TERMINAL_MODEL in gbl_GetConnectionStatus
Global Const gbl_UTS_20 = 0
Global Const gbl_UTS_40 = 1
Global Const gbl_UTS_60 = 2

' Return codes for gbl_CONNECTION_STATUS in gbl_GetConnectionStatus
Global Const gbl_NORMAL = 0
Global Const gbl_BROKEN = 1

' Return codes for gbl_TRANSMIT_MODE in gbl_GetConnectionStatus
Global Const gbl_TRANSMIT_ALLMODE = 0
Global Const gbl_TRANSMIT_VARMODE = 1
Global Const gbl_TRANSMIT_CHANMODE = 2

' Return codes for gbl_KEYBOARD_MODE in gbl_GetConnectionStatus
Global Const gbl_KEYBOARD_UNLOCK = 0
Global Const gbl_KEYBOARD_LOCKED = 1

' Return codes for gbl_FORMS in gbl_GetConnectionStatus
Global Const gbl_FORMS_OFF = 0
Global Const gbl_FORMS_ON = 1

' Return codes for gbl_XMT in gbl_GetConnectionStatus
Global Const gbl_XMT_OFF = 0
Global Const gbl_XMT_ON = 1

' Return codes for gbl_RCV in gbl_GetConnectionStatus
Global Const gbl_RCV_OFF = 0
Global Const gbl_RCV_ON = 1

' Return codes for gbl_LTAI in gbl_GetConnectionStatus
Global Const gbl_LTAI_OFF = 0
Global Const gbl_LTAI_ON = 1

' Global Constants for Character Case
Global Const gbl_UPPER = 1
Global Const gbl_MIXED = 2
  
' Global constants for gbl_StartKeystrokeIntercept
Global Const gbl_AIDKeys = 1
Global Const gbl_AllKeys = 2
Global Const HLL_INTERCEPTAIDKEYS = 1
Global Const HLL_INTERCEPTALLKEYS = 2

' Session Constants
Global Const SESSION_A = 0
Global Const SESSION_B = 1
Global Const SESSION_C = 2
Global Const SESSION_D = 3
Global Const SESSION_E = 4
Global Const SESSION_F = 5
Global Const SESSION_G = 6
Global Const SESSION_H = 7
Global Const SESSION_I = 8
Global Const SESSION_J = 9
Global Const SESSION_K = 10
Global Const SESSION_L = 11
Global Const SESSION_M = 12
Global Const SESSION_N = 13
Global Const SESSION_O = 14
Global Const SESSION_P = 15
Global Const SESSION_Q = 16
Global Const SESSION_R = 17
Global Const SESSION_S = 18
Global Const SESSION_T = 19
Global Const SESSION_U = 20
Global Const SESSION_V = 21
Global Const SESSION_W = 22
Global Const SESSION_X = 23
Global Const SESSION_Y = 24
Global Const SESSION_Z = 25

' Global Constants for gbl_StartHostNotify
Global Const gbl_NOTIFYCURSOR = 1
Global Const gbl_NOTIFYOIA = 2
Global Const gbl_NOTIFYPS = 4
Global Const gbl_NOTIFYBEEPS = 8
Global Const gbl_NOTIFYBASECOLOR = 16
Global Const gbl_NOTIFYMODEL = 32
Global Const gbl_NOTIFYPOWER = 64
Global Const gbl_NOTIFYALL = 127

' Global Constants for gbl_FileTransferDlg
Global Const gbl_FTTSO = 1
Global Const gbl_FTCMS = 2
Global Const gbl_FTCICS = 4
Global Const gbl_FTRECEIVEDLG = 100
Global Const gbl_FTSENDDLG = 200

' Connectivity types for use with the gbl_RegisterClient function
Global Const gbl_3270 = 1
Global Const gbl_5250 = 2
Global Const gbl_ASYNC = 3
Global Const gbl_INFO = 5
Global Const gbl_HP = 6
