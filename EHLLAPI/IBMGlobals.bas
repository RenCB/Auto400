Attribute VB_Name = "IBMGlobals"
'
' IBMGLOBALS  1.0  07/18/2002
'
'  Access Client Solutions Utility Bridge
'                   (C) COPYRIGHT IBM CORP. 2002, 2016
'
' All rights reserved.  Provided on an "AS IS" basis, no
' warranty expressed or implied.
'

Option Explicit
' debug via msgbox
Global blnDebug As Boolean

' debug via file
Global blnWriteToFile As Boolean
' VB trace file
Global strUtlVBtrace As String

Global ConnectSession As String
Global HllFunctionNo As Long
Global HllData       As String * 8000
Global HllLength     As Long
Global HllReturnCode As Long
Global HllParm5 As Long
Global HllParm6 As Long
Global lngTempRC As Long

Declare Function utlapi& Lib "PCSUTL32.DLL" (Func&, ByVal DataString$, Length&, RetC&, Xtra1&, Xtra2&)
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long


'/*********************************************************************/
'/**************** UTLAPI FUNCTION NUMBERS ****************************/
'/*********************************************************************/

Global Const UA_ADD_WAIT = 201                       ' Add Wait
Global Const UA_ADD_WAIT_FOR_CURSOR = 202            ' Add Wait for Cursor
Global Const UA_ADD_WAIT_FOR_CURSOR_MOVE = 203       ' Add Wait for Cursor Move
Global Const UA_ADD_WAIT_FOR_HOST_CONNECT = 204      ' Add Wait for Host connect
Global Const UA_ADD_WAIT_FOR_HOST_DISCONNECT = 205   ' Add Wait for Host Disconnect
Global Const UA_ADD_WAIT_FOR_KEY = 206               ' Add Wait for Key
Global Const UA_ADD_WAIT_FOR_STRING = 207            ' Add Wait for String
Global Const UA_ADD_WAIT_FOR_STRING_NOT_AT = 208     ' Add Wait for String Not at
Global Const UA_ADD_WAIT_HOST_QUIET = 209            ' Add Wait for Host Quiet
Global Const UA_ALLOW_UPDATES = 210                  ' Allow updates
                                                         
Global Const UA_BLOCK_UPDATES = 211                  ' Block updates
Global Const UA_CLEAR_EVENT_TABLE = 212              ' Clear the Event Table
Global Const UA_CLOSE_CONFIGURATION = 213            ' Close the Configuration
Global Const UA_CONNECT_SESSION = 214                ' Connect to a Session
Global Const UA_DELETE_EVENT = 215                   ' Delete an Event
Global Const UA_DISCONNECT_SESSION = 216             ' Disconnect from a Session
Global Const UA_EXECUTE = 217                        ' Execute
Global Const UA_GET_UA_API_VERSION = 218             ' Get the version of UTL APIs
Global Const UA_GET_CONFIGURATION = 219              ' Get the configuration
Global Const UA_GET_CONNECTION_STATUS = 220          ' Get the Status of the connection

Global Const UA_GET_CURSOR_LOCATION = 221            ' Get the Location of the cursor
Global Const UA_GET_EMULATOR_PATH = 222              ' Get the Emulator Path
Global Const UA_GET_EMULATOR_VERSION = 223           ' Get the version of the Emulator
Global Const UA_GET_ERROR = 224                      ' Get error
Global Const UA_GET_FIELD_INFO = 225                 ' Get Field Information
Global Const UA_GET_FIELD_LENGTH = 226               ' Get the Length of the Host field
Global Const UA_GET_FIELD_POSITION = 227             ' Get the position of the Host field
Global Const UA_GET_KEY_STROKE = 228                 ' Get the Key Stroke
Global Const UA_GET_LAYOUT_NAME = 229                ' Get the name of the Layout
Global Const UA_GET_PARAMETER = 230                  ' Get the parameter

Global Const UA_GET_SESSION_HANDLE = 231             ' Get the Session Handle
Global Const UA_GET_SESSIONS = 232                   ' Get the sessions
Global Const UA_GET_SESSION_SIZE = 233               ' Get the Session size
Global Const UA_GET_SESSION_STATUS = 234             ' Get the status of the session
Global Const UA_GET_STRING = 235                     ' Get the String
Global Const UA_GET_STRING_FROM_FIELD = 236          ' Get the String from the field
Global Const UA_HOLD_HOST = 237                      ' Hold the Host
Global Const UA_LIST_SESSIONS = 238                  ' List the Sessions
Global Const UA_LOCK_KEYBOARD = 239                  ' Lock the keyboard
Global Const UA_OPEN_CONFIGURATION = 240             ' Open the configuration

Global Const UA_OPEN_LAYOUT = 241                    ' Open the layout
Global Const UA_PAUSE = 242                          ' A Pause
Global Const UA_RECEIVE_FILE = 243                   ' Receive a CMS/MVS file
Global Const UA_REGISTER_CLIENT = 244                ' Register a client
Global Const UA_RESET_SYSTEM = 245                   ' Reset the system
Global Const UA_RESUME_HOST = 246                    ' Resume the host
Global Const UA_ROW_COLUMN = 247                     ' Row and column
Global Const UA_RUN_EMULATOR_MACRO = 248             ' Run a macro
Global Const UA_RUN_EXTRA_MACRO = 249                ' Run an EXTRA macro
Global Const UA_RUN_EXTRA_MACRO_ASYNC = 250          ' Run an EXTRA macro asynchronously
                                                         
Global Const UA_SEARCH_FIELD = 251                   ' Search for a field
Global Const UA_SEARCH_SESSION = 252                 ' Search for a session
Global Const UA_SEND_AND_WAIT = 253                  ' Send and wait
Global Const UA_SEND_FILE = 254                      ' Send a file
Global Const UA_SEND_KEY = 255                       ' Send a key to the host
Global Const UA_SEND_STRING = 256                    ' Send a string
Global Const UA_SEND_STRING_TO_FIELD = 257           ' Send a string to a field
Global Const UA_SESSION_OFF = 258                    ' Session is off
Global Const UA_SESSION_ON = 259                     ' Session is on
Global Const UA_SET_CURSOR_LOCATION = 260            ' Set the cursor location

Global Const UA_SET_PARAMETER = 261                  ' Set the parameter
Global Const UA_SHOW_LAST_ERROR = 262                ' Show the last error
Global Const UA_START_KEYSTROKE_INTERCEPT = 263      ' Start intercepting those keystrokes
Global Const UA_START_SESSION = 264                  ' Start the Session
Global Const UA_STOP_KEYSTROKE_INTERCEPT = 265       ' Stop intercepting keystrokes
Global Const UA_STOP_SESSION = 266                   ' Stop the session
Global Const UA_UNLOCK_KEYBOARD = 267                ' Unlock the keyboard
Global Const UA_UNREGISTER_CLIENT = 268              ' Unregister client
Global Const UA_WAIT = 269                           ' Wait
Global Const UA_WAIT_FOR_CURSOR = 270                ' Wait for cursor

Global Const UA_WAIT_FOR_CURSOR_MOVE = 271           ' Wait for cursor move
Global Const UA_WAIT_FOR_EVENT = 272                 ' Wait for an event
Global Const UA_WAIT_FOR_HOST_CONNECT = 273          ' Wait for host connect
Global Const UA_WAIT_FOR_HOST_DISCONNECT = 274       ' Wait for host disconnect
Global Const UA_WAIT_FOR_KEY = 275                   ' Wait for a key
Global Const UA_WAIT_FOR_STRING = 276                ' Wait for string
Global Const UA_WAIT_HOST_QUIET = 277                ' Wait for host quiet
Global Const UA_GET_WINDOW_HANDLE = 278              ' Get window handle
Global Const UA_SET_TRACE = 290                      ' For internal use

Public Function printTrace(str As String)
  If blnDebug Then MsgBox str
  If blnWriteToFile Then WriteToFile str
End Function

Public Function printTraceAll()
  If blnDebug Then MsgBox (HllFunctionNo & " " & HllLength & " " & HllReturnCode & " " & HllParm5 & " " & HllParm6 & " " & HllData)
  If blnWriteToFile Then WriteToFile (HllFunctionNo & " " & HllLength & " " & HllReturnCode & " " & HllParm5 & " " & HllParm6 & " " & Mid(HllData, 1, HllLength))
End Function

Public Function printTrace5()
  If blnDebug Then MsgBox (HllFunctionNo & " " & HllLength & " " & HllReturnCode & " " & HllParm5 & " " & HllParm6)
  If blnWriteToFile Then WriteToFile (HllFunctionNo & " " & HllLength & " " & HllReturnCode & " " & HllParm5 & " " & HllParm6)
End Function

Public Function WriteToFile(str As String)
  Dim FileNumber
  FileNumber = FreeFile   ' Get unused file number.
  strUtlVBtrace = "c:\utlVBtrace.txt"
  Open strUtlVBtrace For Append As #FileNumber   ' Create file name.
  Write #FileNumber, str
  Close #FileNumber   ' Close file.
End Function

Public Function printScreen()
  Dim strScreen As String
  Dim rc As Integer
  rc = utl_GetString(1, 1, strScreen, 1920)
  If blnWriteToFile Then WriteToFile strScreen
End Function
