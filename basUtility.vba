'//=================================================================//
'/|     MODULE:  basUtility
'/|    PURPOSE:  Utility Functions for Access Applications
'/|         BY:  Sean
'/|       DATE:  11/30/96
'/|    HISTORY:  11/30/96    Initial Release
'/|              1/11/97     Added Some New Routines
'/|              7/6/2009    UPDATE WITH SOME NEW CODE
'/|              11/1/2012   UPDATE WITH SOME NEW CODE AND COMBINE
'/|                            with modLibrary from Ravi and Cull
'/|                            unused routines.
'/|              04/02/2015  Update RunQueriesByPrefix and
'/|                            FieldExists by Sean and Anjali
'/|              06/01/2015  Update other stuff
'/|              06/04/2015  More updates
'/|                          Added AuditCruise, GetVersion,GetScalar
'/|                          ExecSQL, ChangeSettingsTablesTestMode
'/|                          FormatFilePath, FormatDate
'/|                          dateGetPriorMonth and dateMonthLastDay
'/|              06/05/2015  Added GetScalar Again STM
'/|              02/20/2016  Added Compact&Repair SubRoutines
'/|                                GetAccessFileVersion SubRoutines
'/|
'/|              04/05/2016  TableExists()  & QueryExists() by Sean
'/|              05/04/2016  FileExists()   & FolderExists()
'/|                             and Move GetSetting() back in here
'/|              05/23/2016  UpdateStatus
'/|              08/12/2016  Added Sub AllCodeToDesktop
'/|              10/07/2016  Added WebText, FileIO and some etc.
'/|                         also now incorp to an .accdb 1st time
'/|                         and CleanPipeDelimTextFile("C:\AAA.TXT")
'/|              10/14/2016  ADD TXTCreateTable & TXTImportFromText
'/|              11/18/2016  Added GetCOPTpath()
'/|              07/24/2019  ADD AttachDSNLessTable
'/|              10/10/2019  Added error handling to RunQueriesByPrefix
'/|              01/20/2021  10+ New functions:  AllCodeToDesktop,
'/|                              AnonymizeMyData,
'/|                                     Brack, GetHaversineMiles,
'/|               ExcelColumnLetter2Number and ExcelColumnNumber2Letter,
'/|                              GetVersion, SAP_GetSession,
'/|                              SAP_RunSAPQueryGeneric,
'/|                              UnwindTable,
'/|                              HASHER TOO!
'/|              04/20/2021  Added DDLer and DDLall AND LinkAll
'/|
'/|
'//====================================================================

Option Compare Database
Option Explicit

'// Variables //
Dim OPENFILENAME As typOPENFILENAME
Global Cntr
Public Const gSUCCESS      As Integer = 0
Public Const gFAIL         As Integer = -1

Private mlLastErr       As Long
Private msLastErr       As String
Private msErrSource     As String
Private msErrDesc       As String

Private moVoice         As Object
Private moNotesDB       As Object
Private moNotesSession  As Object
Private moFSO           As Object
'Private moDAL           As clsDAL
Private msFileOut       As String
Private mnFileHandleOut As Integer
Private mlRows          As Long
Private msFileOutPath   As String
Private m_Company       As String
Private m_AppName       As String
Private Const defCompany As String = "VB and VBA Program Settings"


'// API Function Declarations //
#If Win64 And VBA7 Then
    Declare PtrSafe Function GetActiveWindow Lib "user32" () As Long
    Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
    Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OPENFILENAME As typOPENFILENAME) As Long
    Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
    Declare PtrSafe Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Declare PtrSafe Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As typOSVERSIONINFO) As Long
    Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As Long, Rectangle As typRECT) As Long
    Declare PtrSafe Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare PtrSafe Sub SetWindowPos Lib "user32" (ByVal hWnd&, ByVal hWndInsertAfter&, ByVal X&, ByVal Y&, ByVal cX&, ByVal cY&, ByVal wFlags&)
    Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Declare PtrSafe Function WinOpenFile Lib "Kernel32" Alias "OpenFile" (ByVal szFileName As String, OpenBuff As typOFSTRUCT, ByVal Flag As Integer) As Long
    Declare PtrSafe Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Declare PtrSafe Function CreateProcessA Lib "Kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal blnheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupinfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Declare PtrSafe Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
    Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    Declare PtrSafe Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
    Declare PtrSafe Function GetTickCount& Lib "Kernel32" () 'milliseconds
    Declare PtrSafe Function GetSaveFileNameB Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As typOPENFILENAME) As Long
    Declare PtrSafe Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As TYPE_BROWSEINFO) As Long
    Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    Declare PtrSafe Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Declare PtrSafe Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
    '
    ' Win32 Registry functions for reading/writing to Registry
    '
    Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
    Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    Declare PtrSafe Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
    Declare PtrSafe Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
    Declare PtrSafe Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    Declare PtrSafe Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    Declare PtrSafe Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
#Else
    Declare Function GetActiveWindow Lib "user32" () As Long
    Declare Function GetDesktopWindow Lib "user32" () As Long
    Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (OPENFILENAME As typOPENFILENAME) As Long
    Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As typOSVERSIONINFO) As Long
    Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, Rectangle As typRECT) As Long
    Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Declare Sub SetWindowPos Lib "user32" (ByVal hWnd&, ByVal hWndInsertAfter&, ByVal x&, ByVal Y&, ByVal cX&, ByVal cY&, ByVal wFlags&)
    Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Declare Function WinOpenFile Lib "Kernel32" Alias "OpenFile" (ByVal szFileName As String, OpenBuff As typOFSTRUCT, ByVal Flag As Integer) As Long
    Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Declare Function CreateProcessA Lib "Kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal blnheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupinfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
    Declare Function GetTickCount& Lib "Kernel32" () 'milliseconds
    Declare Function GetSaveFileNameB Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As typOPENFILENAME) As Long
    Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Declare Function SHBrowseForFolder Lib "shell32" (lpbi As TYPE_BROWSEINFO) As Long
    Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
    '
    ' Win32 Registry functions for reading/writing to Registry
    '
    Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
    Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Any) As Long
    Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
    Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
#End If


'// Constants //
Global Const OFN_READONLY = &H1                '// OpenFileName Dialog Constants //
Global Const OFN_OVERWRITEPROMPT = &H2
Global Const OFN_HIDEREADONLY = &H4
Global Const OFN_NOCHANGEDIR = &H8
Global Const OFN_SHOWHELP = &H10
Global Const OFN_ENABLEHOOK = &H20
Global Const OFN_ENABLETEMPLATE = &H40
Global Const OFN_ENABLETEMPLATEHANDLE = &H80
Global Const OFN_NOVALIDATE = &H100
Global Const OFN_ALLOWMULTISELECT = &H200
Global Const OFN_EXTENSIONDIFFERENT = &H400
Global Const OFN_PATHMUSTEXIST = &H800
Global Const OFN_FILEMUSTEXIST = &H1000
Global Const OFN_CREATEPROMPT = &H2000
Global Const OFN_SHAREAWARE = &H4000
Global Const OFN_NOREADONLYRETURN = &H8000
Global Const OFN_NOTESTFILECREATE = &H10000
Global Const OFN_SHAREFALLTHROUGH = 2
Global Const OFN_SHARENOWARN = 1
Global Const OFN_SHAREWARN = 0

Global Const OF_EXIST = &H4000   '// WinOpenFile Constant

Global Const EC_NORMAL_PRIORITY_CLASS = &H20&    '// Exec Command
Global Const EC_INFINITE = -1&

Const SM_CXSCREEN = 0            '// Width of screen         '// Get System Metrics Constants //
Const SM_CYSCREEN = 1            '// Height of screen
Const SM_CXFULLSCREEN = 16       '// Width of window client area
Const SM_CYFULLSCREEN = 17       '// Height of window client area
Const SM_CYMENU = 15             '// Height of menu
Const SM_CYCAPTION = 4           '// Height of caption or title
Const SM_CXFRAME = 32            '// Width of window frame
Const SM_CYFRAME = 33            '// Height of window frame
Const SM_CXHSCROLL = 21          '// Width of arrow bitmap on horizontal scroll bar
Const SM_CYHSCROLL = 3           '// Height of arrow bitmap on horizontal scroll bar
Const SM_CXVSCROLL = 2           '// Width of arrow bitmap on vertical scroll bar
Const SM_CYVSCROLL = 20          '// Height of arrow bitmap on vertical scroll bar
Const SM_CXSIZE = 30             '// Width of bitmaps in title bar
Const SM_CYSIZE = 31             '// Height of bitmaps in title bar
Const SM_CXCURSOR = 13           '// Width of cursor
Const SM_CYCURSOR = 14           '// Height of cursor
Const SM_CXBORDER = 5            '// Width of window frame that cannot be sized
Const SM_CYBORDER = 6            '// Height of window frame that cannot be sized
Const SM_CXDOUBLECLICK = 36      '// Width of rectangle around the location of the first click. The
                                 '   second click must occur in the same rectangular location.
Const SM_CYDOUBLECLICK = 37      '// Height of rectangle around the location of the first click. The
                                 '   second click must occur in the same rectangular location.
Const SM_CXDLGFRAME = 7          '// Width of dialog frame window
Const SM_CYDLGFRAME = 8          '// Height of dialog frame window
Const SM_CXICON = 11             '// Width of icon
Const SM_CYICON = 12             '// Height of icon
Const SM_CXICONSPACING = 38      '// Width of rectangles the system uses to position tiled icons
Const SM_CYICONSPACING = 39      '// Height of rectangles the system uses to position tiled icons
Const SM_CXMIN = 28              '// Minimum width of window
Const SM_CYMIN = 29              '// Minimum height of window
Const SM_CXMINTRACK = 34         '// Minimum tracking width of window
Const SM_CYMINTRACK = 35         '// Minimum tracking height of window
Const SM_CXHTHUMB = 10           '// Width of scroll box (thumb) on horizontal scroll bar
Const SM_CYVTHUMB = 9            '// Width of scroll box (thumb) on  vertical scroll bar
Const SM_DBCSENABLED = 42        '// Returns a non-zero if the current Windows version uses double-byte
                                 '   characters, otherwise returns zero
Const SM_DEBUG = 22              '// Returns non-zero if the Windows version is a debugging version
Const SM_MENUDROPALIGNMENT = 40  '// Alignment of popup menus. If zero, left side is aligned with
                                 '   corresponding left side of menu-bar item. If non-zero, left
                                 '   side is aligned with right side of corresponding menu bar item
Const SM_MOUSEPRESENT = 19       '// Non-zero if mouse hardware is installed
Const SM_PENWINDOWS = 41         '// Handle of Pen Windows dynamic link library
                                 '   if Pen Windows is installed
Const SM_SWAPBUTTON = 23         '// Non-zero if the left and right mouse buttons are swapped

Const SM_SYSTEM_RESOURCES = &H0
Const SM_USER_RESOURCES = &H2
Const SM_ENHANCED_MODE = &H20
Const SM_WF_80x87 = &H400

Global Const SW_SHOWNORMAL = 1                '// Show Window Constants //
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3

Public Const VER_PLATFORM_WIN32s = 0          '// OS Version Info Constants //
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Global Const KEY_LBUTTON = &H1                '// Key Codes //
Global Const KEY_RBUTTON = &H2
Global Const KEY_CANCEL = &H3
Global Const KEY_MBUTTON = &H4                '// NOT contiguous with L & RBUTTON //
Global Const KEY_BACK = &H8
Global Const KEY_TAB = &H9
Global Const KEY_CLEAR = &HC
Global Const KEY_RETURN = &HD
Global Const KEY_SHIFT = &H10
Global Const KEY_CONTROL = &H11
Global Const KEY_MENU = &H12
Global Const KEY_PAUSE = &H13
Global Const KEY_CAPITAL = &H14
Global Const KEY_ESCAPE = &H1B
Global Const KEY_SPACE = &H20
Global Const KEY_PRIOR = &H21
Global Const KEY_NEXT = &H22
Global Const KEY_END = &H23
Global Const KEY_HOME = &H24
Global Const KEY_LEFT = &H25
Global Const KEY_UP = &H26
Global Const KEY_RIGHT = &H27
Global Const KEY_DOWN = &H28
Global Const KEY_SELECT = &H29
Global Const KEY_PRINT = &H2A
Global Const KEY_EXECUTE = &H2B
Global Const KEY_SNAPSHOT = &H2C
Global Const KEY_INSERT = &H2D
Global Const KEY_DELETE = &H2E
Global Const KEY_HELP = &H2F
'// KEY_A thru KEY_Z are the same as their ASCII equivalents: 'A' thru 'Z'
'// KEY_0 thru KEY_9 are the same as their ASCII equivalents: '0' thru '9'
Global Const KEY_NUMPAD0 = &H60
Global Const KEY_NUMPAD1 = &H61
Global Const KEY_NUMPAD2 = &H62
Global Const KEY_NUMPAD3 = &H63
Global Const KEY_NUMPAD4 = &H64
Global Const KEY_NUMPAD5 = &H65
Global Const KEY_NUMPAD6 = &H66
Global Const KEY_NUMPAD7 = &H67
Global Const KEY_NUMPAD8 = &H68
Global Const KEY_NUMPAD9 = &H69
Global Const KEY_MULTIPLY = &H6A
Global Const KEY_ADD = &H6B
Global Const KEY_SEPARATOR = &H6C
Global Const KEY_SUBTRACT = &H6D
Global Const KEY_DECIMAL = &H6E
Global Const KEY_DIVIDE = &H6F
Global Const KEY_F1 = &H70
Global Const KEY_F2 = &H71
Global Const KEY_F3 = &H72
Global Const KEY_F4 = &H73
Global Const KEY_F5 = &H74
Global Const KEY_F6 = &H75
Global Const KEY_F7 = &H76
Global Const KEY_F8 = &H77
Global Const KEY_F9 = &H78
Global Const KEY_F10 = &H79
Global Const KEY_F11 = &H7A
Global Const KEY_F12 = &H7B
Global Const KEY_F13 = &H7C
Global Const KEY_F14 = &H7D
Global Const KEY_F15 = &H7E
Global Const KEY_F16 = &H7F
Global Const KEY_NUMLOCK = &H90
    
Global Const OLE_CREATE_EMBED = 0          '// OLE Control Actions //
Global Const OLE_CREATE_NEW = 0            '// for VB compatibility //
Global Const OLE_CREATE_LINK = 1
Global Const OLE_CREATE_FROM_FILE = 1      '// for VB compatibility //
Global Const OLE_COPY = 4
Global Const OLE_PASTE = 5
Global Const OLE_UPDATE = 6
Global Const OLE_ACTIVATE = 7
Global Const OLE_CLOSE = 9
Global Const OLE_DELETE = 10
Global Const OLE_INSERT_OBJ_DLG = 14
Global Const OLE_PASTE_SPECIAL_DLG = 15
Global Const OLE_FETCH_VERBS = 17

Global Const OLE_LINKED = 0                '// OLEType //
Global Const OLE_EMBEDDED = 1
Global Const OLE_NONE = 3

Global Const OLE_EITHER = 2                '// OLETypeAllowed //

Global Const OLE_AUTOMATIC = 0             '// UpdateOptions //
Global Const OLE_FROZEN = 1
Global Const OLE_MANUAL = 2

Global Const OLE_ACTIVATE_MANUAL = 0       '// AutoActivate modes //
Global Const OLE_ACTIVATE_GETFOCUS = 1
Global Const OLE_ACTIVATE_DOUBLECLICK = 2

Global Const OLE_SIZE_CLIP = 0             '// SizeModes //
Global Const OLE_SIZE_STRETCH = 1
Global Const OLE_SIZE_AUTOSIZE = 2
Global Const OLE_SIZE_ZOOM = 3

Global Const OLE_DISPLAY_CONTENT = 0       '// DisplayTypes //
Global Const OLE_DISPLAY_ICON = 1

Global Const VERB_PRIMARY = 0              '// Special Verb Values //
Global Const VERB_SHOW = -1
Global Const VERB_OPEN = -2
Global Const VERB_HIDE = -3
Global Const VERB_INPLACEUIACTIVATE = -4
Global Const VERB_INPLACEACTIVATE = -5

Global Const PROP_CAT_NA = 0            '// Catagory property bitmask for Property object //
Global Const PROP_CAT_LAYOUT = 1
Global Const PROP_CAT_DATA = 2
Global Const PROP_CAT_EVENT = 4
Global Const PROP_CAT_OTHER = 8

Global Const SPECIALEFFECT_NORMAL = 0      '// For <Control>.SpecialEffect ... //
Global Const SPECIALEFFECT_RAISED = 1
Global Const SPECIALEFFECT_SUNKEN = 2

Global Const gALLOWUPDATING_DEFAULT_TABLES = 0   '// For <Form>.AllowUpdating ... //
Global Const gALLOWUPDATING_NO_TABLES = 2

Global Const gDEFAULTEDITING_ALLOW_EDITS = 2  '// For <Form>.DefaultEditing ... //
Global Const gDEFAULTEDITING_READ_ONLY = 3

Global Const TOOLBAR_SHOW_NODATA = "N"        '// View Mode for ToolBar and ALL Forms ... //
Global Const TOOLBAR_SHOW_DATA = "D"
Global Const TOOLBAR_SET_DISPLAY_MODE = "S"
Global Const TOOLBAR_SET_EDIT_MODE = "E"
Global Const TOOLBAR_SET_OUTPUT_MODE = "O"
Global Const TOOLBAR_UNLOAD = "U"

Global Const HWND_TOP = 0                     '// Moves MS Access window to top of Z-order. //

Global Const SWP_NOZORDER = &H4               '// Ignores the hWndInsertAfter. //

Global Const gEDIT_MODE = "E"
Global Const gDISPLAY_MODE = "D"
Global Const gOUTPUT_MODE = "O"
Global Const gSHUTDOWN_MODE = "S"             '// See ToolBar_Unload ...

Global Const SHOW_DATA = "D"
Global Const SHOW_NO_DATA = "N"
 
Global Const gDISPLAY_MODE_BACKCOLOR = 12632256         '// Grey
Global Const gEDIT_MODE_BACKCOLOR = 16777215            '// White

Global Const MB_OK = 0                  '// MsgBox Constants ... //
Global Const MB_OKCANCEL = 1
Global Const MB_YESNOCANCEL = 3
Global Const MB_YESNO = 4
Global Const MB_RETRYCANCEL = 5
Global Const MB_ICONSTOP = 16
Global Const MB_ICONQUESTION = 32
Global Const MB_ICONEXCLAMATION = 48
Global Const MB_ICONINFORMATION = 64
Global Const MB_DEFBUTTON2 = 256
Global Const IDOK = 1
Global Const IDCANCEL = 2
Global Const IDRETRY = 4
Global Const IDYES = 6
Global Const IDNO = 7

Private Const BIF_RETURNONLYFSDIRS = 1    '//RKP/5-10-01. Used by BrowseForFolder.
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Public Const NULL_INTEGER = -32768        '//Alias Null Values for each VB data type
Public Const NULL_LONG = -2147483648#
Public Const NULL_SINGLE = -3.402823E+38
Public Const NULL_DOUBLE = -1.7976931348623E+308
Public Const NULL_CURRENCY = -922337203685477#
Public Const NULL_STRING = ""
Public Const NULL_DATE = #1/1/100#
Public Const NULL_BYTE = 0

Private Const SE_ERR_FNF = 2&          '//Used by ShellExecute() API Function in LaunchAssociatedFile
Private Const SE_ERR_PNF = 3&
Private Const SE_ERR_ACCESSDENIED = 5&
Private Const SE_ERR_OOM = 8&
Private Const SE_ERR_DLLNOTFOUND = 32&
Private Const SE_ERR_SHARE = 26&
Private Const SE_ERR_ASSOCINCOMPLETE = 27&
Private Const SE_ERR_DDETIMEOUT = 28&
Private Const SE_ERR_DDEFAIL = 29&
Private Const SE_ERR_DDEBUSY = 30&
Private Const SE_ERR_NOASSOC = 31&
Private Const ERROR_BAD_FORMAT = 11&

Global Const FONTWEIGHT_NORMAL = 400       '// .FontWeight Constants ... //
Global Const FONTWEIGHT_BOLD = 700

Global Const HEIGHT_POINT_3_5 = 0.35 * 1440
Global Const HEIGHT_POINT_3_0 = 0.3 * 1440
Global Const HEIGHT_POINT_2_5 = 0.25 * 1440

Global Const VARTYPE_EMPTY = 0             '// VarType Constants ... //
Global Const VARTYPE_NULL = 1
Global Const VARTYPE_INTEGER = 2
Global Const VARTYPE_LONG = 3
Global Const VARTYPE_SINGLE = 4
Global Const VARTYPE_DOUBLE = 5
Global Const VARTYPE_CURRENCY = 6
Global Const VARTYPE_DATE = 7
Global Const VARTYPE_STRING = 8

Global Const BLACK = 0                  '// .Fore/BackColor Constants ... //
Global Const BRIGHT_BLUE = 16776960
Global Const BROWN = 33023
Global Const DARK_BLUE = 8388608
Global Const DARK_GREEN = 16384
Global Const GREEN = 8454016
Global Const GREY = 12632256
Global Const GUNMETAL_BLUE = 8421440
Global Const LIGHT_BLUE = 16777088
Global Const MAROON = 8388863
Global Const MEDIUM_GREEN = 32768
Global Const PINK = 16711935
Global Const PURPLE = 16711808
Global Const RED = 255
Global Const WHITE = 16777215
Global Const YELLOW = 8454143

Global Const OBJ_SENSE_MAX = 1
Global Const OBJ_SENSE_MIN = 2

'Constants to manipulate Windows Registry
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_CURRENT_CONFIG = &H80000004 'RKP/v4.0.2/01-11-02. TODO:Not sure about this value. Need to verify...
Private Const ERROR_SUCCESS = 0
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_ARENA_TRASHED = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const KEY_QUERY_VALUE = &H1
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const READ_CONTROL = &H20000
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                       

Private Const HKEY_PERFORMANCE_DATA = &H80000004  ' Constants for Windows 32-bit Registry API
Private Const HKEY_DYN_DATA = &H80000006
Private Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
Private Const REG_OPENED_EXISTING_KEY = &H2                      ' Existing Key opened
Private Const REG_OPTION_RESERVED = 0           ' Parameter is reserved
Private Const REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
Private Const REG_OPTION_CREATE_LINK = 2        ' Created key is a symbolic link
Private Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore
Private Const DELETE = &H10000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
'Private Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const ERROR_MORE_DATA = 234


'// User Types //
Type typOFSTRUCT
   cBytes As String * 1
   fFixedDisk As String * 1
   nErrCode As Integer
   szReserved As String * 4
   szPath As String * 128
End Type

Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lptitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXXountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Type typOPENFILENAME
   lStructSize As Long
   hWndOwner As Long
   hInstance As Long
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Type typOSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128   ' Maintenance string for PSS usage.
End Type

Public Enum fileSizeEnum
   Bytes
   KB
   MB
   GB
   TB
End Enum

Type typRECT
   x1 As Long
   y1 As Long
   x2 As Long
   y2 As Long
End Type

'RKP/5-10-01 - Used by sBrowseForFolder
Public Type TYPE_BROWSEINFO
   hWndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Private Type TYPE_EMAIL_BODY
   sName    As String
   sValue   As String
End Type

Public Type TYPE_KEY_VALUE_PAIR
   Key      As String
   value    As String
End Type


Public Function AddTrustedLocation()
On Error GoTo err_proc
'WARNING:  THIS CODE MODIFIES THE REGISTRY
'sets registry key for 'trusted location'

  Dim intLocns As Integer
  Dim i As Integer
  Dim intNotUsed As Integer
  Dim strLnKey As String
  Dim reg As Object
  Dim strPath As String
  Dim strTitle As String
  
  strTitle = "Add Trusted Location"
  Set reg = CreateObject("wscript.shell")
  strPath = CurrentProject.Path

  'Specify the registry trusted locations path for the version of Access used
  strLnKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Format(Application.Version, "##,##0.0") & _
             "\Access\Security\Trusted Locations\Location"

On Error GoTo err_proc0
  'find top of range of trusted locations references in registry
  For i = 999 To 0 Step -1
      reg.RegRead strLnKey & i & "\Path"
      GoTo chckRegPths        'Reg.RegRead successful, location exists > check for path in all locations 0 - i.
checknext:
  Next
  MsgBox "Unexpected Error - No Registry Locations found", vbExclamation
  GoTo exit_proc
  
  
chckRegPths:
'Check if Currentdb path already a trusted location
'reg.RegRead fails before intlocns = i then the registry location is unused and
'will be used for new trusted location if path not already in registy

On Error GoTo err_proc1:
  For intLocns = 1 To i
      reg.RegRead strLnKey & intLocns & "\Path"
      'If Path already in registry -> exit
      If InStr(1, reg.RegRead(strLnKey & intLocns & "\Path"), strPath) = 1 Then GoTo exit_proc
NextLocn:
  Next
  
  If intLocns = 999 Then
      MsgBox "Location count exceeded - unable to write trusted location to registry", vbInformation, strTitle
      GoTo exit_proc
  End If
  'if no unused location found then set new location for path
  If intNotUsed = 0 Then intNotUsed = i + 1
  
'Write Trusted Location regstry key to unused location in registry
On Error GoTo err_proc:
  strLnKey = strLnKey & intNotUsed & "\"
  reg.RegWrite strLnKey & "AllowSubfolders", 1, "REG_DWORD"
  reg.RegWrite strLnKey & "Date", Now(), "REG_SZ"
  reg.RegWrite strLnKey & "Description", Application.CurrentProject.name, "REG_SZ"
  reg.RegWrite strLnKey & "Path", strPath & "\", "REG_SZ"
  
exit_proc:
  Set reg = Nothing
  Exit Function
  
err_proc0:
  Resume checknext
  
err_proc1:
  If intNotUsed = 0 Then intNotUsed = intLocns
  Resume NextLocn

err_proc:
  MsgBox Err.Description, , strTitle
  Resume exit_proc
  
End Function


Sub AllCodeToDesktop()
''The reference for the FileSystemObject Object is Windows Script Host Object Model
''but it not necessary to add the reference for this procedure.
   Dim fs     As Object
   Dim f      As Object
   Dim strMod As String
   Dim mdl    As Object
   Dim i      As Integer
   
   Dim sfolder      As String
   Dim sHostName    As String
   Dim sUserName    As String
   Dim sUserDomain  As String
   Dim sUserProfile As String
   
   sHostName = Environ$("COMPUTERNAME")
   sUserName = Environ$("USERNAME")
   sUserDomain = Environ$("USERDOMAIN")
   sUserProfile = Environ$("USERPROFILE")
   
   Debug.Print sHostName    '// If InStr(sHostName, "SMACDER") > 0 Thenf
   Debug.Print sUserName
   Debug.Print sUserDomain
   Debug.Print sUserProfile
   Set fs = CreateObject("Scripting.FileSystemObject")
   sfolder = sUserProfile & "\Desktop"
   Set f = fs.CreateTextFile(sfolder & "\" & Replace(CurrentProject.name, ".", "") & ".txt")
   Debug.Print sfolder
   
   ''For each component in the project ...
   For Each mdl In VBE.ActiveVBProject.VBComponents
      If InStr(1, mdl.name, "basUtility") = 0 And InStr(1, mdl.name, "basModule") = 0 And InStr(1, mdl.name, "InterfaceCommon") = 0 And InStr(1, mdl.name, "modExcelPivot") = 0 And _
         InStr(1, mdl.name, "InterfaceSourcingLocks") = 0 And InStr(1, mdl.name, "basUtility") = 0 And InStr(1, mdl.name, "basUtility") = 0 Then
         ''using the count of lines ...
         i = VBE.ActiveVBProject.VBComponents(mdl.name).CodeModule.CountOfLines
         ''put the code in a string ...
         If i > 0 Then
            strMod = VBE.ActiveVBProject.VBComponents(mdl.name).CodeModule.Lines(1, i)
         End If
         ''and then write it to a file, first marking the start with
         ''some equal signs and the component name.
         f.WriteLine String(15, "=") & vbCrLf & mdl.name _
             & vbCrLf & String(15, "=") & vbCrLf & strMod
      End If
   Next
   
   ''Close eveything
   f.Close
   Set fs = Nothing
End Sub


Sub AllDebugListEnvironmentVariables()
   'each environment variable in turn
   Dim EnvironmentVariable As String
   'the number of each environment variable
   Dim EnvironmentVariableIndex As Integer
   'get first environment variables
   EnvironmentVariableIndex = 1
   EnvironmentVariable = Environ(EnvironmentVariableIndex)
   'loop over all environment variables till there are no more
   Do Until EnvironmentVariable = ""
   'get next e.v. and print out its value
   Debug.Print EnvironmentVariableIndex, EnvironmentVariable
   'go on to next one
   EnvironmentVariableIndex = EnvironmentVariableIndex + 1
   EnvironmentVariable = Environ(EnvironmentVariableIndex)
   Loop
End Sub


Function AnonymizeMyData(cellInput As Variant) As Variant

   Dim PN       As Long
   Dim i        As Long
   Dim tmp      As Variant
   Dim Arr()    As String
   Dim charCode As Variant
   
   Select Case True
   
   Case VarType(cellInput) = vbString
      tmp = ""
      'Extract the Characters into an Array
      Arr = Split(StrConv(cellInput, vbUnicode), Chr$(0))
      For i = LBound(Arr) To UBound(Arr)
         If Arr(i) <> "" Then
            charCode = Asc(Arr(i))
            'Character is between A to Z
            If charCode >= 97 And charCode <= 122 Then
               Arr(i) = Chr(Int((122 - 97 + 1) * Rnd) + 97)
            ElseIf charCode >= 65 And charCode <= 90 Then
               Arr(i) = Chr(Int((90 - 65 + 1) * Rnd) + 65)
            'Character is between 0 to 9
            ElseIf charCode >= 48 And charCode <= 57 Then
               Arr(i) = Chr(Int((57 - 48 + 1) * Rnd) + 48)
            End If
         End If
      Next i
      
      '***
      AnonymizeMyData = Join(Arr, "")
      '***
      
   Case VarType(cellInput) = vbDate
      tmp = ""
      On Error Resume Next
      'Generate a random number between 1 and 2 to determine whether
      'to add or subtract the random number
      PN = Int((2 - 1 + 1) * Rnd) + 1
    
      'Generate a random number between 0 and 9999 and add or subtract this number into the date
      If PN = 1 Then
         tmp = cellInput + Int(10000 * Rnd)
      Else
         tmp = cellInput - Int(10000 * Rnd)
      End If
    
      'If error is there, generate a random date between 1-Jan-1990 and 31-Dec-2030
      If Err.Number > 0 Then
         tmp = Int((DateValue("1-Jan-1990") - DateValue("31-Dec-2030") + 1) * Rnd) + DateValue("31-Dec-2030")
      End If
      '''On Error GoTo 0
      '***
      AnonymizeMyData = tmp
      '***
    
   Case IsNumeric(cellInput)
   tmp = ""
   'Extract the Characters into an Array
   Arr = Split(StrConv(CStr(cellInput), vbUnicode), Chr$(0))
   'First digit can not be 0
   Arr(LBound(Arr)) = Chr(Int((57 - 49 + 1) * Rnd) + 49)
   For i = LBound(Arr) + 1 To UBound(Arr)
      If Arr(i) <> "" Then
         charCode = Asc(Arr(i))
         'Character is between 0 to 9
         If charCode >= 48 And charCode <= 57 Then
            Arr(i) = Chr(Int((57 - 48 + 1) * Rnd) + 48)
         End If
      End If
   Next i
   AnonymizeMyData = CDbl(Join(Arr, ""))

   End Select

End Function


Function Brack(strIN As String) As String
   
   Brack = "[" & strIN & "]"

End Function



Function AttachDSNLessTable(stLocalTableName As String, stRemoteTableName As String, stServer As String, stDatabase As String, Optional stUsername As String, Optional stPassword As String)
'//Name     :   AttachDSNLessTable
'//Purpose  :   Create a linked table to SQL Server without using a DSN
'//Parameters
'//     stLocalTableName: Name of the table that you are creating in the current database
'//     stRemoteTableName: Name of the table that you are linking to on the SQL Server database
'//     stServer: Name of the SQL Server that you are linking to
'//     stDatabase: Name of the SQL Server database that you are linking to
'//     stUsername: Name of the SQL Server user who can connect to SQL Server, leave blank to use a Trusted Connection
'//     stPassword: SQL Server user password
    
  'AttachDSNLessTable("tblFredGSanford","dbo.tblSQLFreddy","ESSO2SQLBOOBOOPROD102","DBNAMEYO")  'USERNM;  PWD
      
    On Error GoTo AttachDSNLessTable_Err
    Dim td As TableDef
    Dim stConnect As String
    
    For Each td In CurrentDb.TableDefs
        If td.name = stLocalTableName Then
            CurrentDb.TableDefs.DELETE stLocalTableName
        End If
    Next
      
    If Len(stUsername) = 0 Then
        '//Use trusted authentication if stUsername is not supplied.
        stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & stServer & ";DATABASE=" & stDatabase & ";Trusted_Connection=Yes"
    Else
        '//WARNING: This will save the username and the password with the linked table information.
        stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & stServer & ";DATABASE=" & stDatabase & ";UID=" & stUsername & ";PWD=" & stPassword
    End If
    Set td = CurrentDb.CreateTableDef(stLocalTableName, dbAttachSavePWD, stRemoteTableName, stConnect)
    CurrentDb.TableDefs.Append td
    AttachDSNLessTable = True
    Exit Function

AttachDSNLessTable_Err:
    
    AttachDSNLessTable = False
    MsgBox "AttachDSNLessTable encountered an unexpected error: " & Err.Description

End Function


Function AttachExclusive(AttachedName As String, SourceDB As String, SourceTable As String) As Integer
'//================================================================================//
'/|   FUNCTION: AttachExclusive                                                    |/
'/| PARAMETERS: AttachedName   Name that the attached table will be called         |/
'/|             SourceDB       The path and .MDB of where the table is             |/
'/|             SourceTable    The name of the actual table                        |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|    PURPOSE: Attach an Access table exclusively.                                |/
'/|      USAGE: AttachExclusive("BLUE","C:\DDI\SAW_DATA.MDB","Species")            |/
'/|         BY: Sean                                                               |/
'/|       DATE: 11/30/96                                                           |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim db As Database, td As TableDef
On Error GoTo AttachExclusive_Err
  
   AttachExclusive = False
   Set db = CurrentDb
   Set td = db.CreateTableDef(AttachedName)
   td.SourceTableName = SourceTable
   td.Connect = ";DATABASE=" & SourceDB & ";"
   td.Attributes = DB_ATTACHEXCLUSIVE
   db.TableDefs.Append td
   AttachExclusive = True

AttachExclusive_Done:
  Exit Function

AttachExclusive_Err:
  AttachExclusive = False
  Resume AttachExclusive_Done
End Function


Function Audit(strTableName As String)
    Dim db As Database
    Dim td As TableDef
    Dim qd As QueryDef
    Dim q2 As QueryDef
    Dim fd As Field
    Dim i As Integer
    Dim j As Integer
    Dim strSQL As String

    Audit = False
    
    Set db = CurrentDb
    
    'TABLE LOOP
    For i = 0 To db.TableDefs.count - 1
        Set td = db.TableDefs(i)
        If td.name = strTableName Then
            Debug.Print "Table " & td.name & " is in " & db.name
            For j = 0 To db.QueryDefs.count - 1
                Set qd = db.QueryDefs(j)
                strSQL = qd.sql
                If InStr(strSQL, strTableName) Then
                   Debug.Print "    " & qd.name
                End If
            Next j
        End If
    Next i
    
    'QUERY LOOP
    For i = 0 To db.QueryDefs.count - 1
        Set q2 = db.QueryDefs(i)
        If q2.name = strTableName Then
            Debug.Print "Query " & td.name & " is in " & db.name
            For j = 0 To db.QueryDefs.count - 1
                Set qd = db.QueryDefs(j)
                strSQL = qd.sql
                If InStr(strSQL, strTableName) Then
                   Debug.Print "    " & qd.name
                End If
            Next j
        End If
    Next i
    
    Audit = True
End Function


Function Audit2(strTableName As String)
    Dim db As Database
    Dim td As TableDef
    Dim qd As QueryDef
    Dim q2 As QueryDef
    Dim fd As Field
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strSQL As String

    Audit2 = False
    
    Set db = CurrentDb
    
    'TABLE LOOP
    For i = 0 To db.TableDefs.count - 1
        Set td = db.TableDefs(i)
        If td.name = strTableName Then
            Debug.Print "Table " & td.name & " is in " & db.name
            For j = 0 To db.QueryDefs.count - 1
                Set qd = db.QueryDefs(j)
                strSQL = qd.sql
                If InStr(strSQL, strTableName) Then
                   Debug.Print "    " & qd.name
                End If
            Next j
        End If
    Next i
    
    'QUERY LOOP
    For i = 0 To db.QueryDefs.count - 1
        Set q2 = db.QueryDefs(i)
           For j = 0 To db.QueryDefs.count - 1
                Set qd = db.QueryDefs(j)
                strSQL = qd.sql
                If InStr(strSQL, strTableName) Then
                   Debug.Print "    " & qd.name
                End If
            Next j
        'field loop
        'For K = 0 To q2.Fields.Count - 1
        '    Set fd = q2.Fields(K)
        '    strSQL = fd.Name
        '    If InStr(strSQL, strTableName) Then
        '        Debug.Print "        " & q2.Name & "   " & fd.Name
        '    End If
        'Next K
    Next i
    
    Audit2 = True
End Function


Function AuditCruise()
   Dim DefaultWorkspace As Workspace
   Dim MyDatabase As Database
   Dim MyTableDef As TableDef
   Dim MyQueryDef As QueryDef
   Dim MyField As Field
   Dim i As Integer
   Dim j As Integer
   Dim rs As dao.Recordset
   
   Set DefaultWorkspace = DBEngine.Workspaces(0)
   Set MyDatabase = DefaultWorkspace.Databases(0)
   Set rs = CurrentDb.OpenRecordset("SELECT * FROM temptblChanges")
   rs.MoveLast
   rs.MoveFirst
    
   While rs.EOF = False
      DoEvents
      Set MyTableDef = MyDatabase.TableDefs(rs!Tbl)
      j = Audit(MyTableDef.name)
      rs.MoveNext
   Wend
End Function


Public Function CallWebService(ByVal vsWebServer As String, ByRef rbSuccess As Boolean, ByVal vsTaskID As String, ByVal vsMillID As String, ByVal vsParamArray As String) As ADODB.Recordset
'**********************************************
'Author  :  Ravi Poluri
'Date/Ver:  04-08-09/V01
'Input   :
'Output  :
'Comments:
'**********************************************
   On Error GoTo Err_Handler
   
   Dim loRecordsets  As Collection
   Dim lsUrl         As String
    
   rbSuccess = True
'   mlLastErr = 0
   
   'msWebServer = Application.ThisWorkbook.VBProject.HelpFile
   'If msWebServer = "" Or VBA.Trim(msWebServer) = "???" Then
   'msWebServer = Sheets("Main").Range("WebServer").Value
   'End If
   
   'lsUrl = "http://" & msWebServer & "/bmos/CallWebService.asmx" & "?" & UrlEncode("task_id") & "=" & UrlEncode(vsTaskID) & "?" & UrlEncode("mill_id") & "=" & UrlEncode(vsMillID) & "?" & UrlEncode("param_array") & "=" & UrlEncode(vsParamArray)
   lsUrl = "http://" & vsWebServer & "/CallWebService.asmx" & "?" & UrlEncode("task_id") & "=" & UrlEncode(vsTaskID) & "?" & UrlEncode("mill_id") & "=" & UrlEncode(vsMillID) & "?" & UrlEncode("param_array") & "=" & UrlEncode(vsParamArray)
   Debug.Print lsUrl
   
   'get the Recordset Collection
   Set loRecordsets = ConvertXmlDocumentToRecordsetsCollection(InvokeLiteWebService(lsUrl))
   
   If loRecordsets Is Nothing Then
      rbSuccess = False
'      mlLastErr = -1
   Else
      If loRecordsets(1).BOF = True And loRecordsets(1).EOF = True Then
         rbSuccess = False
      Else
         Set CallWebService = loRecordsets(1)
      End If
   End If
   
   'Set CallWebService = loRecordsets
   
Err_Handler:
   If Err Then
'      MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Public Function ChangeSettingsTablesTestMode(sType As String) As Boolean
'***********************************************************************
'Author  :  STM
'Date/Ver:  11-04-13/V01
'USAGE   :  i = ChangeSettingsTablesTestMode("TEST")
'           i = ChangeSettingsTablesTestMode("LIVE")
'Input   :
'Output  :
'Comments:
'***********************************************************************
Dim sTimeStamp As String
   sTimeStamp = VBA.Format(VBA.Now(), "yyyymmdd-hhmm")
   DoCmd.Rename "tbl000Settings" & "_bkup_" & sTimeStamp, acTable, "tbl000Settings"
   DoCmd.Rename "tbl000SettingsExcelPivots" & "_bkup_" & sTimeStamp, acTable, "tbl000SettingsExcelPivots"

   Select Case sType
   Case "TEST"
      DoCmd.CopyObject , "tbl000Settings", acTable, "tbl000Settings_Test"
      DoCmd.CopyObject , "tbl000SettingsExcelPivots", acTable, "tbl000SettingsExcelPivots_Test"
   Case "LIVE"
      DoCmd.CopyObject , "tbl000Settings", acTable, "tbl000Settings_Live"
      DoCmd.CopyObject , "tbl000SettingsExcelPivots", acTable, "tbl000SettingsExcelPivots_Live"
   End Select
Debug.Print sTimeStamp
ChangeSettingsTablesTestMode = True
End Function


Function ChangeSQL(strQueryName As String, strSQLText As String) As Integer
'//================================================================================//
'/|   FUNCTION: ChangeSQL                                                          |/
'/| PARAMETERS: strQueryName, a String identifying the Query to Process            |/
'/|             strSQLText,   a String containing the new SQL string to assign     |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|    PURPOSE: Change the SQL String of a Query                                   |/
'/|      USAGE: ChangeSQL("Query1","SELECT DISTINCTROW Examples.* FROM Examples;") |/
'/|         BY: Sean                                                               |/
'/|       DATE: 11/30/96                                                           |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim db As Database, qd As QueryDef
On Error GoTo ChangeSQL_Err
    
    ChangeSQL = False                       '// Default - Failed
    
    Set db = DBEngine(0)(0)
    Set qd = db.QueryDefs(strQueryName)     '// Open the QueryDef Object
    qd.sql = strSQLText                     '// Assign the Query a New SQL String
    ChangeSQL = True                        '// Indicate a Successful Completion

ChangeSQL_Exit:
    On Error Resume Next                    '// Trap if there was an error opening the QueryDef
    qd.Close                                '// Close the QueryDef
    Exit Function

ChangeSQL_Err:
    Resume ChangeSQL_Exit

End Function


Function CIntlNumber(strNumber As String, iDecPlaces As Integer) As Double
'//================================================================================//
'/|   FUNCTION: CIntlNumber                                                        |/
'/| PARAMETERS: strNumber,  a String identifying the NUMBER to Process             |/
'/|             iDecPlaces, integer, number of decimal places to use               |/
'/|    RETURNS: Number in Double Precision                                         |/
'/|    PURPOSE: Convert International Number in String format to local machine Dbl |/
'/|      USAGE: CIntlNumber("755.6389",5)          CIntlNumber("57",5)             |/
'/|             CIntlNumber("-6.578296",2)                                         |/
'/|         BY: Sean                                                               |/
'/|       DATE: 8/14/2012                                                          |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim sdbg As String
Dim ddec As Double
Dim intSign As Integer
On Error GoTo CIntlNumber_Err
    
    CIntlNumber = 0             '// Default - Failed
    intSign = 1
    strNumber = Trim(strNumber)
    'CIntlNumber = strNumber
    If Left(strNumber, 1) = "-" Then intSign = -1
    
    If InStr(1, strNumber, ".") > 0 Then
        CIntlNumber = Left(strNumber, InStr(1, strNumber, ".") - 1)
    Else
        CIntlNumber = CDbl(strNumber)
    End If
    
    sdbg = PadRight(Right(strNumber, Len(strNumber) - InStr(1, strNumber, ".")), "0", iDecPlaces)
     'Debug.Print "...       " & sdbg
    ddec = 1 * (CLng(sdbg) / (10 ^ iDecPlaces))
     'Debug.Print "...       " & ddec
    
    If InStr(1, strNumber, ".") > 0 Then CIntlNumber = CIntlNumber + (ddec * intSign)
    '// Successful Completion

CIntlNumber_Exit:
    On Error Resume Next                    '// Trap if there was an error opening the QueryDef
    Exit Function

CIntlNumber_Err:
    Resume CIntlNumber_Exit

End Function


Sub CleanPipeDelimTextFile(sFilePath As String)
   'CleanPipeDelimTextFile("C:\Users\smacder.NAIPAPER\Desktop\Text File Proc\D--PulpForecastHdr123PIPEDeLim_2MB.txt")
   
   ' Manipulating a text file with VBA
   ' Loops through text file and creates revised one
   ' This code requires a reference (Tools > References) to Microsoft Scripting Runtime
   
   Dim FSO As FileSystemObject   'needs ms scripting runtime reference
   Dim FSOFile As TextStream, FSOFileRevised As TextStream
   Dim FilePath As String, FilePathRevised As String
   Dim iStart As Integer
   Dim iCounter As Integer
   Dim sLine As String
   Dim iLineCounter As Integer
   Dim iSpaces As Integer
   
   FilePath = sFilePath
   
   FilePathRevised = Left(FilePath, Len(FilePath) - 4) & "_CLN" & Right(FilePath, 4)
   
   Set FSO = New FileSystemObject
   If FSO.FileExists(FilePath) Then
      ' opens the file for reading
      Set FSOFile = FSO.OpenTextFile(FilePath, 1, False)
      ' opens "revised" file in write mode
      Set FSOFileRevised = FSO.OpenTextFile(FilePathRevised, 2, True)
      iStart = 0
   
      Do While Not FSOFile.AtEndOfStream
         sLine = FSOFile.ReadLine
         iStart = 1
         iLineCounter = 1
         
         sLine = Trim(sLine)
         sLine = ReplaceString(sLine, "&ndash;?", "-")
         
         If iStart >= 1 And sLine <> "" And InStr(1, sLine, "|") >= 1 And _
               Left(Trim(sLine), 51) <> "---------------------------------------------------" And _
               Left(Trim(sLine), 51) <> "|--------------------------------------------------" Then 'Contains PIPE
            ' write maniplulation code here
            If sLine = "Date      Time    School" Then
               sLine = "Dt        Tm      SCHOOL"
            End If
            
            If Left(Trim(sLine), 51) = "---------------------------------------------------" Then  'KILL IT
               sLine = ""
            End If
            
            If Left(Trim(sLine), 51) = "|--------------------------------------------------" Then  'KILL IT
               sLine = ""
            End If
            
            If InStr(1, sLine, "Summary by PO Item") > 0 Then   'KILL IT
               sLine = ""
            End If
            
            If InStr(1, sLine, "---------------------------------------------------") > 0 Then   'KILL IT
               sLine = ""
            End If
            
            If Left(sLine, 1) = "|" Then sLine = Right(sLine, Len(sLine) - 1)  'IF LINE BEGINS WITH "|" ELIMINATE IT
            If Left(sLine, 2) = " |" Then sLine = Right(sLine, Len(sLine) - 2) 'IF LINE BEGINS WITH "|" ELIMINATE IT
            If Right(sLine, 1) = "|" Then sLine = Left(sLine, Len(sLine) - 1)  'IF LINE ENDS   WITH "|" ELIMINATE IT
            
            If iSpaces < 0 Then iSpaces = 1
            FSOFileRevised.WriteLine (sLine)                         'writes line as is
            
            ''FSOFileRevised.Write Trim(sLine) & Space(iSpaces) & "  |  "    'writes but does not Carriage Return Line Feed
            ''FSOFileRevised.WriteLine Trim(sLine)                         'writes line as is
         End If
            
         If InStr(1, sLine, "**exit**") > 0 Then Exit Do                                   'LAST LINE
         iLineCounter = iLineCounter + 1
      Loop
        
      FSOFile.Close
      FSOFileRevised.Close
    Else
      MsgBox (FilePath & " does not exist")
    End If
    
    'Debug.Print "done"
End Sub


Sub CleanTabDelimTextFile(sFilePath As String)
   'CleanTabDelimTextFile("C:\Users\smacder.NAIPAPER\Desktop\Text File Proc\aaa.txt")
   
   ' Manipulating a text file with VBA
   ' Loops through text file and creates revised one
   ' This code requires a reference (Tools > References) to Microsoft Scripting Runtime
   
   Dim FSO As FileSystemObject   'needs ms scripting runtime reference
   Dim FSOFile As TextStream, FSOFileRevised As TextStream
   Dim FilePath As String, FilePathRevised As String
   Dim iStart As Integer
   Dim iCounter As Integer
   Dim sLine As String
   Dim iLineCounter As Integer
   Dim iSpaces As Integer
   
   FilePath = sFilePath
   
   FilePathRevised = Left(FilePath, Len(FilePath) - 4) & "_CLN" & Right(FilePath, 4)
   
   Set FSO = New FileSystemObject
   If FSO.FileExists(FilePath) Then
      ' opens the file for reading
      Set FSOFile = FSO.OpenTextFile(FilePath, 1, False)
      ' opens "revised" file in write mode
      Set FSOFileRevised = FSO.OpenTextFile(FilePathRevised, 2, True)
      iStart = 0
   
      Do While Not FSOFile.AtEndOfStream
         sLine = FSOFile.ReadLine
         iStart = 1
         iLineCounter = 1
         
         sLine = Trim(sLine)
         sLine = ReplaceString(sLine, "&ndash;?", "-")
         
         If iStart >= 1 And sLine <> "" And Right(Trim(sLine), 21) <> "List contains no data" And _
                     InStr(1, sLine, "ZDP_CONS_ACTUAL") < 1 Then 'Good Line
            ' write maniplulation code here
            ' --NONE--

            FSOFileRevised.WriteLine (sLine)                         'writes line as is
            
            ''FSOFileRevised.Write Trim(sLine) & Space(iSpaces) & "  |  "    'writes but does not Carriage Return Line Feed
            ''FSOFileRevised.WriteLine Trim(sLine)                         'writes line as is
         End If
            
         If InStr(1, sLine, "**exit**") > 0 Then Exit Do                                   'LAST LINE
         iLineCounter = iLineCounter + 1
      Loop
        
      FSOFile.Close
      FSOFileRevised.Close
    Else
      MsgBox (FilePath & " does not exist")
    End If
    
    'Debug.Print "done"
End Sub


Sub CloseObj(intObjType As Integer, strObjName As String)
'//=============================================================================//
'/|        SUB:  CloseObj                                                       |/
'/| PARAMETERS:  intObjType, Type of Object to Close (A_FORM or A_REPORT)       |/
'/|              strObjName, Name of Object to Close                            |/
'/|    RETURNS:  -NONE-                                                         |/
'/|    PURPOSE:  Closes a Form or Report without error trapping                 |/
'/|      USAGE:  Call CloseObj(A_FORM, "Form1")                                 |/
'/|         BY:  Sean                                                           |/
'/|       DATE:  11/30/96                                                       |/
'/|    HISTORY:                                                                 |/
'//=============================================================================//

On Error Resume Next

    DoCmd.Close intObjType, strObjName

End Sub


Public Function CommonDialog( _
   ByVal vlFileDialogType As Office.MsoFileDialogType, _
   ByVal vsFileFilters As String, _
   ByVal vnFileIndex As Integer, _
   ByVal vsTitle As String, _
   ByVal vsButtonText As String, _
   ByVal vbMultiSelect As Boolean) As String
'**********************************************
'Author  :  Ravi Poluri
'Date    :  01-21-03/v6.8.110
'Input   :
'Output  :  e.g. "C:\DATA\Image\colors.xls"
'Comments:
'vsFileFilters = "Text Files,*.txt;*.csv|All Files,*.*"
'msoFileDialogFilePicker = 3
'basUtility.CommonDialog(msoFileDialogFilePicker, "Text Files,*.txt;*.csv|Excel Files,*.xls|All Files,*.*", 2, "Select a file", "Select a file", False)
'**********************************************

   Dim fd         As FileDialog
   Dim file       As String
   Dim filters()  As String
   Dim ctr        As Integer
   Dim filterDesc As String
   Dim filterExt  As String

   'Set fd = Application.FileDialog(msoFileDialogFilePicker)
   Set fd = Application.FileDialog(vlFileDialogType)
   
   'fd.Filters.Add "All Files", "*.*" 'Default
   'fd.filters.Add "Text Files", "*.txt;*.csv"
   ctr = 0
   If VBA.InStr(1, vsFileFilters, "|", vbTextCompare) > 0 Then
      filters = VBA.Split(vsFileFilters, "|", , vbTextCompare)
      For ctr = 0 To UBound(filters)
         DoEvents
         filterDesc = VBA.Split(filters(ctr), ",", , vbTextCompare)(0)
         filterExt = VBA.Split(filters(ctr), ",", , vbTextCompare)(1)
         fd.filters.Add filterDesc, filterExt
      Next
   Else
      ReDim filters(ctr)
      filters(ctr) = vsFileFilters
      filterDesc = VBA.Split(filters(ctr), ",", , vbTextCompare)(0)
      filterExt = VBA.Split(filters(ctr), ",", , vbTextCompare)(1)
      fd.filters.Add filterDesc, filterExt
   End If
   
   fd.FilterIndex = vnFileIndex '2
   file = fd.Show
   If fd.SelectedItems.count > 0 Then
      CommonDialog = fd.SelectedItems(1)
   Else
      CommonDialog = ""
   End If

Err_Handler:
   If Err Then
'      ProcessMsg Err.Number, Err.Description, "", "DownloadFile"
      MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Public Function CommonDialogSaveAs() As String
   Dim vFile As String
   vFile = GetSaveFilename()
   'If vFile <> "" Then MsgBox vFile
   CommonDialogSaveAs = vFile
End Function


Sub CompactRepairCurrentDB(Optional sVersion As String = "2016")
'USAGE:
'  Call CompactRepairCurrentDB("2016")
'  Call CompactRepairCurrentDB("2007")
'
   Dim AltKey As String
   Dim CtrlKey As String
   Dim ShiftKey As String
   Dim TabKey As String
   Dim EnterKey As String
   '--------------------------
   AltKey = "%"
   CtrlKey = "^"
   ShiftKey = "+"
   TabKey = "{TAB}"
   EnterKey = "~"
   '--------------------------
   
   Select Case sVersion
   
      Case "2016", "2013", "2010"
         ''Access 2016/2013/2010: SendKeys ALT-F  |  TAB  |  TAB  |  ENTER
         ''===============================================================
         'Alt|File Menu
         SendKeys AltKey & "(F)", False
         WaitSecs (1) 'Application.Wait Now + TimeValue("00:00:01")  'Wait One Second
         
         'Tab Tab to get to Compact & Repair
         SendKeys TabKey, False
         SendKeys TabKey, False
      
         'Enter to Do it.
         SendKeys EnterKey, False

      Case "2007"
         ''Access 2007: SendKeys ALT-F  |  M  |  C
         ''===============================================================
         'Alt|File|Manage|Compact&Repair
         SendKeys AltKey & "(F)", False
         WaitSecs (1)  'Wait One Second
         SendKeys AltKey & "(M)", False
         WaitSecs (1)  'Wait One Second
         SendKeys AltKey & "(C)", False
      
      Case "2003" Or "2002"
         ''Access 2003/2002: SendKeys ALT-T  |  D  |  C
         ''===============================================================
         'Alt|Tools|Database|Compact&Repair
         SendKeys AltKey & "(T)", False
         WaitSecs (1)  'Wait One Second
         SendKeys AltKey & "(D)", False
         WaitSecs (1)  'Wait One Second
         SendKeys AltKey & "(C)", False
      
      Case Else
         MsgBox "Please update your software to a version of Access that was released this Century.  "

   End Select

End Sub


Sub CompactRepairExternalDB(pathToMDB)
'USAGE
'   CompactRepairExternalDB "C:\OPTMODELS\PC31\PC31APP.MDB"
'   CompactRepairExternalDB "C:\OPTMODELS\PC31\CAPDATA.MDB"
'   CompactRepairExternalDB "C:\OPTMODELS\PC31\MTXDATA.MDB"
'   CompactRepairExternalDB "C:\OPTMODELS\TS1\TS1.MDB"

   Dim objScript    As Object
   Dim objAccess    As Object
   Dim strPathToMDB As String
   Dim strMsg       As String
   Dim strTempDB    As String

   strPathToMDB = pathToMDB '"C:\Work\Temp\SoutheastTours.mdb"
   
   ' Set a name and path for a temporary mdb file
   strTempDB = Left(strPathToMDB, Len(strPathToMDB) - 3) & "bak" '"C:\Work\Temp\SoutheastTours.bak"

   ' Create Access Application Object
   Set objAccess = CreateObject("Access.Application")

   ' Perform the DB Compact into the temp mdb file
   ' (If there is a problem, then the original mdb is  preserved)
   objAccess.DBEngine.CompactDatabase strPathToMDB, strTempDB

   If Err.Number > 0 Then
      ' There was an error.  Inform the user and halt execution
      strMsg = "The following error was encountered while compacting database:"
      strMsg = strMsg & vbCrLf & vbCrLf & Err.Description
   Else
      ' Create File System Object to handle file manipulations
      Set objScript = CreateObject("Scripting.FileSystemObject")

      ' Back up the original file as Filename.mdbz.  In case of undetermined
      ' error, it can be recovered by simply removing the terminating "z".
      'objScript.CopyFile strPathToMDB, strPathToMDB & "z", True

      ' Copy the compacted mdb by into the original file name
      objScript.CopyFile strTempDB, strPathToMDB, True

      ' We are finished with TempDB.  Kill it.
      objScript.DeleteFile strTempDB
   End If

   ' Always remember to clean up after yourself
   Set objAccess = Nothing
   Set objScript = Nothing
   '

   MsgBox "Finished...  " & strPathToMDB, , "COMPACT AND REPAIR"
End Sub


Public Function GetSaveFilename(Optional ByVal vFileFilter As String, Optional ByVal _
   vWindowTitle As String, Optional ByVal vInitialDir As String, Optional ByVal _
   vInitialFileName As String, Optional ByVal vMultiSelect As Boolean) As String
   Dim ofn As typOPENFILENAME
   Dim RetVal As Long
   ofn.lStructSize = Len(ofn)
   ofn.hWndOwner = 0
   ofn.hInstance = 0
   ofn.lpstrFile = VBA.IIf(vInitialDir = "", VBA.Space$(254), vInitialDir)
   ofn.lpstrInitialDir = VBA.IIf(vWindowTitle = "", VBA.CurDir, vInitialDir)
   ofn.lpstrTitle = VBA.IIf(vWindowTitle = "", "Select File", vWindowTitle)
   ofn.lpstrFilter = VBA.IIf(vFileFilter = "", "All Files (*.*)" & VBA.Chr(0) & "*.*", VBA.Replace(vFileFilter, ",", VBA.Chr$(0)))
   ofn.nMaxFile = 255
   ofn.lpstrFileTitle = VBA.Space$(254)
   ofn.nMaxFileTitle = 255
   ofn.flags = 0
   RetVal = GetSaveFileNameB(ofn)
   If RetVal Then GetSaveFilename = VBA.Left(VBA.Trim$(ofn.lpstrFile), VBA.Len(VBA.Trim$(ofn.lpstrFile)) - 1)
End Function


Public Function config_GetValue(ByVal vsConfigFile As String, ByVal vsKeyName As String) As String
'**********************************************
'Author  :  RKP
'Date/Ver:  09-22-08/V01
'Input   :
'Output  :
'Comments:
'http://www.perl.com/pub/a/2001/04/17/msxml.html

'The following is how a *.config file looks like:

'<configuration>
'   <appSettings>
'      <!--COMMENTS-->
'      <!--Mon,22Sep08-->
'      <!--EMailProcessor.config-->
'      <!--
'    The key "processEMailOnStartup" is used by the file:
'    EMailProcessor_V01.xls
'    or
'    EMailProcessor_VXX.xls
'    (where XX refers to a version number)
'    to determine whether (or not) to "refresh" (process email) on startup (as defined in the ThisWorkbook.Workbook_Open event).
'    False - Do not run the "refresh" on startup
'    True - Run the "refresh" on startup
'    The default value is "False".
'    The VBScript file, "LaunchExcel.vbs" is the only place where the value of the key "processEMailOnStartup" is set to "True".
'    The "LaunchExcel.vbs" file sets the key to "True" at the beginning of the script and sets the value to "False" at the end of the script.
'    The file "LaunchExcel.vbs" is launched by a Schedule Task (Control Panel) called "LaunchExcel".
'    The sole purpose of the "LaunchExcel.vbs" file is to launch the Excel file:
'    EMailProcessor_VXX.xls
'    The "LaunchExcel.vbs" file is scheduled to run every day at 7:30 am.
'    A log file (ScheduledTask.log) is updated every time the "refresh" task is performed.
'    -->
'      <add key="processEMailOnStartup" value="True"/>
'      <add key="processCurrentDayFilesOnly" value="True"/>
'      <!--
'      RKP/2009-05-31
'      This switch will just process email.
'      The goal of this switch is to allow email to be processed multiple times a day
'      without processing data for the OTIS report each time.
'      -->
'      <add key="processEMailOnly" value="True"/>
'      <add key="deleteEmailAfterProcessing" value="False"/>
'      <add key="fileLastUpdateTimeStamp" value="6/25/2009 3:56:08 PM"/>
'      <!--
'      The following keys are required for email processing.
'      The first key "emailCount" tells the processor how many email subjects are available to be processed.
'      The next sequence of "emailSubject_XX" keys are the actual email subjects that are to be processed.
'      -->
'      <add key="emailCount" value="10"/>
'      <add key="emailSubject_01" value="BW AHQ - MTS - Review Shippable Quantities Report"/>
'      <add key="emailSubject_02" value="BW AHQ - MTO - Shipping Capacity Report"/>
'      <add key="emailSubject_03" value="% Sales Query for Missing Reason Codes Report"/>
'      <add key="emailSubject_04" value="% Stops Query for Missing Reason Codes Report"/>
'      <add key="emailSubject_05" value="Complaint Mgmt - Invoice Register"/>
'      <add key="emailSubject_06" value="OTIS Report - Port to Port 1st Leg Shipments"/>
'      <add key="emailSubject_07" value="Daily Production Stats - V1 and V4"/>
'      <add key="emailSubject_08" value="OTIS - % Stops - AHQ_RKP_LO_DEL_O13_Q002_V001"/>
'      <add key="emailSubject_09" value="OTIS - % Sales - AHQ_RKP_LO_DEL_O13_Q002_V002"/>
'      <add key="emailSubject_10" value="OTIS - % Sales - AHQ_RKP_LO_DEL_C08_Q002_V001"/>
'   </appSettings>
'</configuration>
'**********************************************
   On Error GoTo Err_Handler

'   Dim sFile      As String
   Dim oDOMDoc    As Object 'MSXML2.DOMDocument30
   Dim oDOMList   As Object ' MSXML2.IXMLDOMNodeList
   Dim oDOMNode   As Object ' MSXML2.IXMLDOMNode
   Dim oDOMEle    As Object ' MSXML2.IXMLDOMElement
   Dim oFSO       As Object ' Scripting.FileSystemObject
   
   'Set oFSO = New Scripting.FileSystemObject
   Set oFSO = CreateObject("Scripting.FileSystemObject")
   
   If oFSO.FileExists(vsConfigFile) = False Then GoTo Err_Handler
   
   'RKP/10-15-08
   'For some strange reason, Application.ThisWorkbook.Path is returning "".
   'sFile = Application.ThisWorkbook.Path & "\EMailProcessor.config"
   
'   vsConfigFile = oFSO.GetParentFolderName(Application.ThisWorkbook.FullName) & "\EMailProcessor.config"
   
'   Debug.Print VBA.Dir(vsConfigFile)
   'Set oDOMDoc = New MSXML2.DOMDocument30
   Set oDOMDoc = CreateObject("Msxml2.DOMDocument.3.0")
   oDOMDoc.async = False
   oDOMDoc.validateOnParse = True
   'oDOMDoc.setProperty "language", "XPath"
   oDOMDoc.Load vsConfigFile
   'Debug.Print oDOMDoc.XML
   'Set oDOMList = oDOMDoc.selectNodes("//appSettings/add[contains(.,'Startup')]")
   Set oDOMEle = oDOMDoc.DocumentElement
   Set oDOMList = oDOMEle.getElementsByTagName("add")
   For Each oDOMNode In oDOMList
      DoEvents
      'Debug.Print oDOMNode.Attributes.getNamedItem("key").Text & ", " & oDOMNode.Attributes.getNamedItem("value").Text
      'If oDOMNode.Attributes.getNamedItem("value").nodeValue = "True" Then
      
'      If oDOMNode.Attributes.getNamedItem("key").Text = "fileLastUpdateTimeStamp" Then
'         oDOMNode.Attributes.getNamedItem("value").nodeValue = VBA.Now()
'      ElseIf oDOMNode.Attributes.getNamedItem("key").Text = "processEMailOnStartup" Then
'         oDOMNode.Attributes.getNamedItem("value").nodeValue = "True"
'      ElseIf oDOMNode.Attributes.getNamedItem("key").Text = "deleteEmailAfterProcessing" Then
'         oDOMNode.Attributes.getNamedItem("value").nodeValue = VBA.Now()
'      End If
      'End If
      'Debug.Print oDOMNode.Attributes.getNamedItem("key").Text & ", " & oDOMNode.Attributes.getNamedItem("value").Text
      If oDOMNode.Attributes.getNamedItem("key").text = vsKeyName Then
         config_GetValue = oDOMNode.Attributes.getNamedItem("value").text
         Exit For
      End If
   Next
   'oDOMDoc.Save Application.ThisWorkbook.Path & "\EMailProcessor.config"
   
   Set oDOMEle = Nothing
   Set oDOMNode = Nothing
   Set oDOMList = Nothing
   Set oDOMDoc = Nothing
   
   
Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", ""
      MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Public Function config_SetValue(ByVal vsConfigFile As String, ByVal vsKeyName As String, ByVal vsKeyValue As String)
'**********************************************
'Author  :  RKP
'Date/Ver:  09-22-08/V01
'Input   :
'Output  :
'Comments:
'http://www.perl.com/pub/a/2001/04/17/msxml.html
'**********************************************
   On Error GoTo Err_Handler

   Dim oDOMDoc    As Object ' MSXML2.DOMDocument30
   Dim oDOMList   As Object ' MSXML2.IXMLDOMNodeList
   Dim oDOMNode   As Object ' MSXML2.IXMLDOMNode
   Dim oDOMEle    As Object ' MSXML2.IXMLDOMElement
   
   Set oDOMDoc = CreateObject("Msxml2.DOMDocument.3.0")
   oDOMDoc.async = False
   oDOMDoc.validateOnParse = True
   oDOMDoc.Load vsConfigFile
   
   Set oDOMEle = oDOMDoc.DocumentElement
   Set oDOMList = oDOMEle.getElementsByTagName("add")
   For Each oDOMNode In oDOMList
      DoEvents
      If oDOMNode.Attributes.getNamedItem("key").text = vsKeyName Then
         oDOMNode.Attributes.getNamedItem("value").text = vsKeyValue
         Exit For
      End If
   Next
   oDOMDoc.Save vsConfigFile 'Application.ThisWorkbook.Path & "\EMailProcessor.config"
   
   Set oDOMEle = Nothing
   Set oDOMNode = Nothing
   Set oDOMList = Nothing
   Set oDOMDoc = Nothing
   
Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", ""
      MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Private Sub ConvertFileEncoding( _
               ByVal FileIn As String, _
               ByVal EncodingIn As String, _
               ByVal LineSeparatorIn As ADODB.LineSeparatorEnum, _
               ByVal FileOut As String, _
               ByVal EncodingOut As String, _
               ByVal LineSeparatorOut As ADODB.LineSeparatorEnum, _
               Optional ByVal NoBOM As Boolean = False, _
               Optional ByVal BOMSize As Integer = 0)
    
    '    ConvertFileEncoding "ascii.txt", "ascii", adCRLF, "utf-16le.txt", "unicode", adCRLF
    '    ConvertFileEncoding "ascii.txt", "ascii", adCRLF, "utf-8-crlf.txt", "utf-8", adCRLF
    '    ConvertFileEncoding "ascii.txt", "ascii", adCRLF, "utf-8-lf.txt", "utf-8", adLF
    '    ConvertFileEncoding "ascii.txt", "ascii", adCRLF, "utf-8-lf-nobom.txt", "utf-8", adLF, True, 3

    'D--PulpForecastHdr123PIPEDeLim_2MBTRY4103.txt
    
    Dim stmIn As ADODB.Stream
    Dim stmOut As ADODB.Stream
    Dim strLine As String
    Dim bytData() As Byte
    
    Set stmIn = New ADODB.Stream
    With stmIn
        .Open
        .Type = adTypeText
        .Charset = EncodingIn
        .LoadFromFile FileIn
        .LineSeparator = LineSeparatorIn
        Set stmOut = New ADODB.Stream
        With stmOut
            .Open
            .Type = adTypeText
            .Charset = EncodingOut
            .LineSeparator = LineSeparatorOut
        End With
        Do Until .EOS
            strLine = .ReadText(adReadLine)
            stmOut.WriteText strLine, adWriteLine
        Loop
        .Close
    End With
    With stmOut
        If NoBOM Then
            .Position = 0 'Must be at 0 to change Type.
            .Type = adTypeBinary
            .Position = BOMSize 'Skip over BOM.
            bytData = .Read(adReadAll)
            'Empty the Stream.
            .Position = 0
            .SetEOS
            .Write bytData
            Erase bytData
            .SaveToFile FileOut, adSaveCreateOverWrite
        Else
            .SaveToFile FileOut, adSaveCreateOverWrite
        End If
    End With
End Sub


Private Function ConvertRecordsetsCollectionToXmlDocument(ByVal voRecordsets As Collection) As Object 'MSXML.DOMDocument
'**********************************************
'Author  :  Ravi Poluri
'Date/Ver:  10-17-05/v2.0
'Input   :
'Output  :
'Comments:
'**********************************************
   On Error GoTo Err_Handler

   Dim loXmlDoc As Object 'New MSXML.DOMDocument
   Dim loXmlResults As Object 'MSXML.IXMLDOMElement
   Dim loRecordset As ADODB.Recordset
   Dim loXmlRecordsetDoc As Object 'New MSXML.DOMDocument
   
   Set loXmlResults = loXmlDoc.createElement("results")
   loXmlDoc.appendChild loXmlResults
   
   For Each loRecordset In voRecordsets
      
      loXmlRecordsetDoc.LoadXML ConvertRecordsetToXml(loRecordset)
      loXmlResults.appendChild loXmlRecordsetDoc.DocumentElement
   Next loRecordset
   
   Set ConvertRecordsetsCollectionToXmlDocument = loXmlDoc

Err_Handler:
   If Err Then
'      MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Private Function ConvertRecordsetToXml(ByVal voRecordset As ADODB.Recordset) As String
'**********************************************
'Author  :  Ravi Poluri, International Paper
'Date/Ver:  10-17-05/v2.0
'Input   :
'Output  :
'Comments:
' These two functions show how to create the XML document representing multiple
' recordsets using VBA code. Note that this code will work in VB.NET with little
' modification.
'**********************************************

    Dim loStream As New ADODB.Stream

    loStream.Open
    voRecordset.Save loStream, adPersistXML
    loStream.Position = 0
    
    ConvertRecordsetToXml = loStream.ReadText()

Err_Handler:
   If Err Then
'      MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Public Function ConvertXmlDocumentToRecordsetsCollection( _
    ByVal voXmlDocument As Object _
) As Collection
'**********************************************
'Author  :  Ravi Poluri
'Date/Ver:  11-28-03/v7.0.113
'Input   :
'ByVal voXmlDocument As MSXML.DOMDocument
'Output  :
'Comments:
'Converts the XML document that was created on the server
'into a collection of Recordsets, representing a DataSet.
'**********************************************
   On Error GoTo Err_Handler

   Dim loRecordsetNodes As Object 'MSXML.IXMLDOMNodeList
   Dim loRecordsetNode As Object 'MSXML.IXMLDOMNode
   Dim loRecordsets As Collection
   
   If Not voXmlDocument Is Nothing Then
   
      Set loRecordsets = New Collection
      
      'loop through recordset nodes
      Set loRecordsetNodes = voXmlDocument.DocumentElement.ChildNodes
      For Each loRecordsetNode In loRecordsetNodes
         'loop through recordset nodes
         loRecordsets.Add ConvertXmlToRecordset(loRecordsetNode.XML)
      Next loRecordsetNode
      
      'return the collection of (disconnected) Recordsets
      Set ConvertXmlDocumentToRecordsetsCollection = loRecordsets
   End If
   
Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", "ConvertXmlDocumentToRecordsetsCollection"
'      MsgBox Err.Number & " - " & Err.Description
      'Resume
   End If
End Function


Private Function ConvertXmlToRecordset(ByVal vsXml As String) As ADODB.Recordset
'**********************************************
'Author  :  Ravi Poluri
'Date/Ver:  11-28-03/v7.0.113
'Input   :
'Output  :
'Comments:
' Converts the XML for a single Recordset back into a
' Recordset.
'**********************************************
   On Error GoTo Err_Handler
   
   Dim loRecordset As New ADODB.Recordset
   Dim loStream As New ADODB.Stream

   loStream.Open
   loStream.WriteText vsXml
   loStream.Position = 0
   loRecordset.Open Source:=loStream, Options:=ADODB.CommandTypeEnum.adCmdFile
       
   Set ConvertXmlToRecordset = loRecordset
   
Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", "ConvertXmlToRecordset"
'      MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Function CopyDir(SourcePath As String, destpath As String) As Integer
'
' Copies files in SourcePath to DestPath
' Does not copy any subdirectories of SourcePath
'
'  Calling Convention:
'    X = CopyDir("C:\Temp", "A:")
'
Dim FileName As String, Copied As Integer
  FileName = Dir(SourcePath & "\*.*")
  Do While FileName <> ""
    Copied = CopyFileX(SourcePath & "\" & FileName, destpath & "\" & FileName)
    If Not Copied Then
      CopyDir = False
      Exit Function
    End If
    FileName = Dir
  Loop
  CopyDir = True
End Function


Function CopyFileX(SourceName As String, DestName As String) As Integer
'
' Copies a single file SourceName to DestName
'
' Calling convention:
'   X = CopyFileX("C:\This.Exe", "C:\That.Exe")
'   X = CopyFileX("C:\This.Exe", "C:\Temp\This.Exe")
'
Const BufSize = 8192
Dim Buffer As String * BufSize, TempBuf As String
Dim SourceF As Integer, DestF As Integer, i As Long
  On Error GoTo CFError
  SourceF = FreeFile
  Open SourceName For Binary As #SourceF
  DestF = FreeFile
  Open DestName For Binary As #DestF
  For i = 1 To LOF(SourceF) \ BufSize
    Get #SourceF, , Buffer
    Put #DestF, , Buffer
  Next i
  i = LOF(SourceF) Mod BufSize
  If i > 0 Then
    Get #SourceF, , Buffer
    TempBuf = Left$(Buffer, i)
    Put #DestF, , TempBuf
  End If
  Close #SourceF
  Close #DestF
  CopyFileX = True

CFExit:
  Exit Function

CFError:
  Close
  MsgBox "Error " & Err.Number & " copying files" & Chr$(13) & Chr$(10) & Error
  CopyFileX = False
  Resume CFExit
End Function


Sub CreateSPTQuery(QName As String, ODBCConnect As String, sql As String)
'//=============================================================================//
'/|        SUB:  CreateSPTQuery                                                 |/
'/| PARAMETERS:  QName         New Query Name as String                         |/
'/|              ODBCConnect   ODBC Connect String                              |/
'/|              SQL           SQL Statement String                             |/
'/|    RETURNS:  -NONE-                                                         |/
'/|    PURPOSE:  Programmatically Create a Pass-Through Query                   |/
'/|      USAGE:  CreateSPTQuery "TestQuery","ODBC;DSN=ncaccess;UID=malcolms;_   |/
'/|              PWD=malcolms;DATABASE=pubs","Select * from dbo.authors"        |/
'/|         BY:  Sean                                                           |/
'/|       DATE:  11/30/96                                                       |/
'/|    HISTORY:                                                                 |/
'//=============================================================================//

Dim db As Database, q As QueryDef
   Set db = CurrentDb
   Set q = db.CreateQueryDef(QName)
   q.Connect = ODBCConnect
   q.sql = sql
   q.Close
   db.Close
End Sub


Function CreateSQLDataSource(SourceToCreate As String, ServerName As String, DatabaseName As String) As Boolean
'//=============================================================================//
'/|   FUNCTION:  CreateSQLDataSource                                            |/
'/| PARAMETERS:  SourceToCreate   Arbitrary name to give the datasource         |/
'/|              ServerName       The name of the server where the database is  |/
'/|              DatabaseName     The name of the database                      |/
'/|    RETURNS:  TRUE = Success; FALSE = Failure                                |/
'/|    PURPOSE:  Programmatically Create a datasource in ODBC.INI.              |/
'/|      USAGE:  CreateSQLDataSource ("MySQLData","DDI_THREE","cboard.DAT"      |/
'/|         BY:  Sean                                                           |/
'/|       DATE:  11/30/96                                                       |/
'/|    HISTORY:                                                                 |/
'//=============================================================================//

Dim Attribs As String, CR As String
  
   CR = Chr$(13)         ' will not work if you use chr$(13) & chr$(10)
   Attribs = "Description=" & SourceToCreate & " Data Source" & CR
   Attribs = Attribs & "OemToAnsi=No" & CR
   Attribs = Attribs & "Server=" & ServerName & CR
   Attribs = Attribs & "Database=" & DatabaseName
   On Error Resume Next
   DBEngine.RegisterDatabase SourceToCreate, "SQL Server", True, Attribs
   
   If Err.Number = 0 Then
      CreateSQLDataSource = True
   Else
      CreateSQLDataSource = False
   End If

End Function


Function CRLF() As String
'//=====================================================//
'/|   FUNCTION:  CRLF                                   |/
'/| PARAMETERS:  -NONE-                                 |/
'/|    RETURNS:  CarriageReturn/LineFeed Pair           |/
'/|    PURPOSE:  Returns a CarriageReturn/LineFeed Pair |/
'/|      USAGE:  Print "GOO" & CRLF() & "GOO"           |/
'/|         BY:  Sean                                   |/
'/|       DATE:  11/30/96                               |/
'/|    HISTORY:                                         |/
'//=====================================================//

    CRLF = Chr$(13) & Chr$(10)

End Function


Public Function date_GetPriorMonth(ByVal currentMonth As Integer) As String
'**********************************************
'Author  :  RKP
'Date/Ver:  08-20-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    If currentMonth = 1 Then
        date_GetPriorMonth = "12"
    Else
        date_GetPriorMonth = VBA.Format(currentMonth - 1, "00")
    End If

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function


Public Function date_MonthLastDay(ByVal dCurrDate As Date)
  Dim dFirstDayNextMonth As Date
  
  On Error GoTo lbl_Error
 
  date_MonthLastDay = Empty
  dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
  date_MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
  
  Exit Function
lbl_Error:
  MsgBox Err.Description, vbOKOnly + vbExclamation
End Function


Function DaysInMonth(dteInput As Date) As Integer
    Dim intDays As Integer

    ' Add one month, subtract dates to find difference.
    intDays = DateSerial(Year(dteInput), Month(dteInput) + 1, Day(dteInput)) _
        - DateSerial(Year(dteInput), Month(dteInput), Day(dteInput))
    DaysInMonth = intDays
End Function




Function DConcatA(ConcatColumns As String, Tbl As String, Optional Criteria As String = "", _
    Optional Delimiter1 As String = ", ", Optional Delimiter2 As String = ", ", _
    Optional Distinct As Boolean = True, Optional Sort As String = "Asc", _
    Optional Limit As Long = 0)
    
    ' Function by Patrick G. Matthews, basically embellishing an approach seen in many
    ' incarnations over the years
    'https://www.experts-exchange.com/articles/2380/Domain-Aggregate-for-Concatenating-Values-by-Group-in-Microsoft-Access.html
    
    ' Requires reference to Microsoft DAO library
    
    ' This function is intended as a "domain aggregate" that concatenates (and delimits) the
    ' various values rather than the more usual Count, Sum, Min, Max, etc.  For example:
    '
    '    Select Field1, DConcatA("Field2", "SomeTable", "[Field1] = '" & Field1 & "'") AS List
    '    FROM SomeTable
    '    GROUP BY Field1
    '
    ' will return the distinct values of Field1, along with a concatenated list of all the
    ' distinct Field2 values associated with each Field1 value.
    
    ' ConcatColumns is a comma-delimited list of columns to be concatenated (typically just
    '   one column, but the function accommodates multiple).  Place field names in square
    '   brackets if they do not meet the customary rules for naming DB objects
    ' Tbl is the table/query the data are pulled from.  Place table name in square brackets
    '   if they do not meet the customary rules for naming DB objects
    ' Criteria (optional) are the criteria to be applied in the grouping.  Be sure to use And
    '   or Or as needed to build the right logic, and to encase text values in single quotes
    '   and dates in #
    ' Delimiter1 (optional) is the delimiter used in the concatenation (default is ", ").
    '   Delimiter1 is applied to each row in the code query's result set
    ' Delimiter2 (optional) is the delimiter used in concatenating each column in the result
    '   set if ConcatColumns specifies more than one column (default is ", ")
    ' Distinct (optional) determines whether the distinct values are concatenated (True,
    '   default), or whether all values are concatenated (and thus may get repeated)
    ' Sort (optional) indicates whether the concatenated string is sorted, and if so, if it is
    '   Asc or Desc.  Note that if ConcatColumns has >1 column and you use Desc, only the last
    '   column gets sorted
    ' Limit (optional) places a limit on how many items are placed into the concatenated string.
    '   The Limit argument works as a TOP N qualifier in the SELECT clause
    
    Dim rs As dao.Recordset
    Dim sql As String
    Dim ThisItem As String
    Dim FieldCounter As Long
    
    On Error GoTo ErrHandler
    
    ' Initialize to Null
    
    DConcatA = Null
    
    ' Build up a query to grab the information needed for the concatenation
    
    sql = "SELECT " & IIf(Distinct, "DISTINCT ", "") & _
            IIf(Limit > 0, "TOP " & Limit & " ", "") & _
            ConcatColumns & " " & _
        "FROM " & Tbl & " " & _
        IIf(Criteria <> "", "WHERE " & Criteria & " ", "") & _
        Switch(Sort = "Asc", "ORDER BY " & ConcatColumns & " Asc", _
            Sort = "Desc", "ORDER BY " & ConcatColumns & " Desc", True, "")
        
    ' Open the recordset and loop through it:
    ' 1) Concatenate each column in each row of the recordset
    ' 2) Concatenate the resulting concatenated rows in the function's return value
    
    Set rs = CurrentDb.OpenRecordset(sql)
    With rs
        Do Until .EOF
            
            ' Initialize variable for this row
            
            ThisItem = ""
            
            ' Concatenate columns on this row
            
            For FieldCounter = 0 To rs.Fields.count - 1
                ThisItem = ThisItem & Delimiter2 & Nz(rs.Fields(FieldCounter).value, "")
            Next
            
            ' Trim leading delimiter
            
            ThisItem = Mid(ThisItem, Len(Delimiter2) + 1)
            
            ' Concatenate row result to function return value
            
            DConcatA = Nz(DConcatA, "") & Delimiter1 & ThisItem
            .MoveNext
        Loop
        .Close
    End With
    
    ' Trim leading delimiter
    
    If Not IsNull(DConcatA) Then DConcatA = Mid(DConcatA, Len(Delimiter1) + 1)
    
    GoTo Cleanup

ErrHandler:
    
    ' Error is most likely an invalid database object name, or bad syntax in the Criteria
    
    DConcatA = CVErr(Err.Number)
    
Cleanup:
    Set rs = Nothing
    
End Function


Public Function DConcatB(strDataSource As String, _
                        strConcatenateField As String, _
                        strDelimiter As String, _
                        ParamArray aFldVal() As Variant) _
       As String
    '*************************************************************************
    'DConcatB()
    'Written by Azli Hassan, http://azlihassan.com/apps
    '© Azli Hassan, http://azlihassan.com/apps
    '
    'Updated (5/6/2018): Can now pass a single WHERE string to aFldVal
    '                    as you would with regular domain aggregate functions.
    'Updated (5/26/2018): Added option to specify string to use as delimiter
    '
    'PURPOSE:   To concatenate all the values of a field in a
    '           table or query that meets the grouping of the
    '           calling query
    '
    'ARGUMENTS:
    ' 1) strDataSource [String]
    '    - Name of table/query that field to be concatenated is in
    '    - May also be an SQL string to be used to set a recordset.
    ' 2) strConcatenateField [String]
    '    - Name of field to be concatenated.
    ' 2) strDelimiter [String]
    '    - String to use as delimiter (seperator) between value.
    ' 4) aFldVal() [Array]
    '    - An array of GroupBy fields and their values.
    '      Must be pass from the query in a repeating order of
    '      Field name (as a string), then the Field value, and so on
    '      By default, concatenated values are ordered ascendingly.
    '      If you need a particular sort order then you'll need to
    '      pass a single WHERE and ORDER BY SQL statement to aFldVal().
    '
    'RETURNS: Concatenated string of UNIQUE values of concatenated string
    '         where data sources fields match the calling queries groupings
    '
    'TIP: Remove or Comments out the Debug.Print statements
    '     AFTER you understand how the function works.
    '*************************************************************************
    ''SELECT tlkMachine.FacilityAbbr, DConcatB("tlkMachine","MachineCode","  @&  ","FacilityAbbr",[FacilityAbbr]) AS MACS
    '' FROM tlkMachine;

    On Error GoTo ErrMsg:

    Dim db As dao.Database, _
        rst As dao.Recordset, _
        fldConcatenate As dao.Field
    Dim strParamArray As String, _
        lngNumOfElements As Long, _
        blnIsNumeric As Boolean
    Dim strSQL As String, _
        strSELECT As String, _
        strFROM As String, _
        strWhere As String, _
        strCRITERIA As String
    Dim blnText As Boolean
    Dim strAdd As String

    'Check that table/query exists in current database
    'IsTableQuery() - http://support.microsoft.com/kb/210398/
    If Not IsTableQuery("", strDataSource) Then
        Dim rstTemp As dao.Recordset
        Set rstTemp = CurrentDb.OpenRecordset(strDataSource)
        If rstTemp.BOF And rstTemp.EOF Then GoTo ExitHere
        rstTemp.Close
        strFROM = "FROM (" & Left(strDataSource, Len(strDataSource) - 1) & _
                  ") as myDataSource "
    Else
        strFROM = "FROM " & strDataSource & " "
    End If

    Set db = CurrentDb()
    Set rst = db.OpenRecordset(strDataSource)

    'Check that table/query has data
    If rst.BOF And rst.EOF Then GoTo ExitHere

    'Check if parramarray is empty
    If IsEmpty(aFldVal) Then
        DConcatB = "#ERR-EmptyParramarray"
    End If
    
    'Check if only 1 thing was passed to parramarray.
    'If so, then assume the whole WHERE string was passed.
    If LBound(aFldVal) = UBound(aFldVal) Then
        'Only 1 element was passed
        strWhere = "WHERE "
        strWhere = strWhere & CStr(aFldVal(LBound(aFldVal)))
        strWhere = strWhere & ";"
    Else
        'More than 1 thing passed
        'Get number of elements in parramarray
        If LBound(aFldVal) = 0 Then
            lngNumOfElements = UBound(aFldVal) + 1
        Else
            lngNumOfElements = UBound(aFldVal) + 1
        End If
        'Check that paramarray has even number of elements
        If lngNumOfElements Mod 2 <> 0 Then Exit Function
        Dim i As Long
        For i = LBound(aFldVal) To UBound(aFldVal) Step 2
            blnIsNumeric = IsNumeric(aFldVal(i + 1))
            Select Case blnIsNumeric
            Case True
                strParamArray = strParamArray & "'" & aFldVal(i) & _
                                "', " & aFldVal(i + 1) & ", "
            Case False
                If IsDate(aFldVal(i + 1)) Then
                strParamArray = strParamArray & "'" & aFldVal(i) & _
                                "', #" & aFldVal(i + 1) & "#, "
                Else
                strParamArray = strParamArray & "'" & aFldVal(i) & _
                                "', '" & aFldVal(i + 1) & "', "
                End If
            End Select
        Next i
        strParamArray = Left(strParamArray, _
                             Len(strParamArray) - Len(", "))
    
        Debug.Print "DConcatB('" & strDataSource & _
                    "', '" & strConcatenateField & "', " & _
                    strParamArray & ")"
    
        For i = LBound(aFldVal) To (UBound(aFldVal)) Step 2
            blnText = (rst.Fields(aFldVal(i)).Type = dbChar) Or _
                      (rst.Fields(aFldVal(i)).Type = dbMemo) Or _
                      (rst.Fields(aFldVal(i)).Type = dbText)
            Select Case blnText
            Case True
                strCRITERIA = "[" & aFldVal(i) & "] = '" & _
                              aFldVal(i + 1) & "'"
            Case False
                If rst.Fields(aFldVal(i)).Type = dbDate Then
                    strCRITERIA = "[" & aFldVal(i) & "] = " & _
                                  "#" & aFldVal(i + 1) & "#"
                Else
                    strCRITERIA = "[" & aFldVal(i) & "] = " & _
                                  aFldVal(i + 1)
                End If
            End Select
            strWhere = strWhere & strCRITERIA & " AND "
        Next i
        strWhere = "WHERE (" & strWhere
        strWhere = Left(strWhere, Len(strWhere) - Len(" AND "))
        strWhere = strWhere & ");"
    End If
    
    
    'Create SQL String to select distinct records
    'that match the query's "GroupBy" values
    'e.g. SELECT DISTINCT Reference
    '     FROM tblData
    '     WHERE ((ProductID=2211) AND (Description="10ï¿½F 15V"));

    strSELECT = "SELECT DISTINCT " & strConcatenateField & " "
    
    strSQL = strSELECT & strFROM & strWhere
    Debug.Print strSQL

    Set rst = db.OpenRecordset(strSQL)

    'Check that SQL recordset has data
    If rst.BOF And rst.EOF Then GoTo ExitHere
    rst.MoveFirst

    'Set recordset field Object
    Set fldConcatenate = rst.Fields(strConcatenateField)

    'Loop through ALL the records
    'in the SELECT DISTICT recordset
    While Not rst.EOF
        With rst
            If DConcatB = "" Then
                'First value
                If Not IsNull(fldConcatenate) Then
                    DConcatB = fldConcatenate
                End If
            Else
                If Not IsNull(fldConcatenate) Then
                    strAdd = strDelimiter & fldConcatenate
                    If InStr(1, DConcatB, _
                             strAdd, vbTextCompare) = 0 Then
                        'Only add if unique
                        DConcatB = DConcatB & strAdd
                    End If
                End If
            End If
        End With
        rst.MoveNext
    Wend

ExitHere:
    On Error Resume Next
    rstTemp.Close
    rst.Close
    db.Close
    Exit Function

ErrMsg:
    DConcatB = "#ERR" & Err.Number
    Debug.Print "Err.Number = " & Err.Number & _
                ", Err.Description = " & Err.Description
    Resume ExitHere

End Function


Sub DDLall(Optional sSQLDB = "BMOS")
'---------------------------------------------------------------------------------------
' Procedure : DDLall
' Author    : Sean
' Purpose   : Generate a SQL Server Script to create all the tables
' Input     : sSQLDB, the SQL Server DB to USE
' Usage     : Call DDLall("FiberOptCentral")  or  Call DDLall()
' Rev       : 2021-04-20
'---------------------------------------------------------------------------------------
    Dim lTbl As Long
    Dim dBase As Database
    Dim Handle As Integer
    Set dBase = CurrentDb

    Handle = FreeFile
    Open GetPathName(GetCurrentMDBwPath) & GetCurrentMDBName & "_DDL.txt" For Output Access Write As #Handle

    For lTbl = 0 To dBase.TableDefs.count - 1
         'If the table name is a temporary or system table then ignore it
        If Left(dBase.TableDefs(lTbl).name, 1) = "~" Or Left(dBase.TableDefs(lTbl).name, 4) = "MSYS" Then
             '~ indicates a temporary table
             'MSYS indicates a system level table
        ElseIf Len(dBase.TableDefs(lTbl).name) > 1 Then
          Print #Handle, DDLer(dBase.TableDefs(lTbl).name)
        Else
        End If
    Next lTbl
    
    Close Handle
    Set dBase = Nothing
    Debug.Print "DDLall Complete."
End Sub


Public Function DDLer(sTableName As String, Optional sSQLDB = "BMOS") As String
'-----------------------------------------------------------------------------------------------
' Procedure : DDLer
' Author    : Sean
' Purpose   : Generate a SQL Server Script to create a table
' Input     : sTableName, the TableName to generate Create Table Statement for
'             sSQLDB,     the SQL Server DB to USE
' Usage     : sSQL = DDLer("tlkFacility")  or sSQL = DDLer("tlkProduct", "FiberOptCentral")
' Rev       : 2021-04-20
'------------------------------------------------------------------------------------------------
   Dim dBase As Database
   Set dBase = CurrentDb
   Dim oTableDef As TableDef
   Dim fldDef As Field
   Dim ctr As Integer
   Dim idxctr As Integer
   Dim fld As Field
   Dim idx As Index
   Dim FieldIndex As Integer
   Dim fldName As String, fldDataInfo As String
   Dim DDL As String
   Dim tableName As String
   Dim sClusterType As String
   Dim sNullType As String
   
   DDLer = ""
   Set oTableDef = dBase.TableDefs(sTableName)
   
   tableName = oTableDef.name
   tableName = Replace(tableName, " ", "_")
   
   DDL = DDL & ""
   DDL = DDL & "USE [" & sSQLDB & "] " & vbCrLf
   DDL = DDL & "GO " & vbCrLf
   DDL = DDL & "" & vbCrLf
   DDL = DDL & "--  Object:  Table [dbo].[" & tableName & "] Script Date: " & Format(Now(), "yyyy-mm-dd") & " " & vbCrLf
   DDL = DDL & "--        :  Generated From " & CurrentDb.name & " " & vbCrLf
   DDL = DDL & "--           FIND THE CODE IN THE DDLer FUNCTION AND DDLall " & vbCrLf
   DDL = DDL & "" & vbCrLf
   DDL = DDL & "DROP TABLE IF EXISTS [dbo].[" & tableName & "]" & vbCrLf
   DDL = DDL & "GO" & vbCrLf
   DDL = DDL & "SET ANSI_NULLS ON" & vbCrLf
   DDL = DDL & "GO" & vbCrLf
   DDL = DDL & "SET QUOTED_IDENTIFIER ON" & vbCrLf
   DDL = DDL & "GO" & vbCrLf
   DDL = DDL & "" & vbCrLf

   DDL = DDL & "CREATE TABLE [dbo].[" & tableName & "] (" & vbCrLf

   With oTableDef
      For FieldIndex = 0 To .Fields.count - 1
         Set fldDef = .Fields(FieldIndex)
   
         With fldDef
            fldName = .name
            fldName = Replace(fldName, " ", "_")
            sNullType = "NULL"
            If .AllowZeroLength = False Then
               sNullType = "NOT NULL"
            End If
            
            Select Case .Type                   ''Double    goes to FLOAT     ''Int, Long goes to INT    ''Text      goes to NVARCHAR
            Case dbBoolean
               fldDataInfo = "BIT " & sNullType & ""
            Case dbDouble, dbByte, dbCurrency
               fldDataInfo = "FLOAT " & sNullType & ""
            Case dbInteger, dbLong, dbSingle
               If .Attributes And dbAutoIncrField Then
                  fldDataInfo = "INT IDENTITY(1,1)"
               Else
                  fldDataInfo = "INT " & sNullType & ""
               End If
            Case dbDate
               fldDataInfo = "DATETIME " & sNullType & ""
            Case dbText
               fldDataInfo = "NVARCHAR(" & Format$(.Size) & ") " & sNullType & ""
            Case dbLongBinary
               fldDataInfo = "****"
            Case dbMemo
               fldDataInfo = "NVARCHAR(255) " & sNullType & ""
            Case dbGUID
               fldDataInfo = "NVARCHAR(16) " & sNullType & ""
            End Select
         End With
      
         If FieldIndex > 0 Then
            DDL = DDL & ", " & vbCrLf
         End If
         
         DDL = DDL & "  " & Brack(fldName) & " " & fldDataInfo
            
      Next FieldIndex
   End With
   
   If oTableDef.Indexes.count > 0 Then
      DDL = DDL & ", " & vbCrLf
   End If
   
   idxctr = 0
   For Each idx In oTableDef.Indexes
      idxctr = idxctr + 1

      If idx.Fields.count > 1 Then
         sClusterType = "CLUSTERED"
      Else
         sClusterType = "NONCLUSTERED"
      End If
      
      If idx.Primary = True Then             '//PRIMARY KEY
         DDL = DDL & "  CONSTRAINT PK_" & oTableDef.name & Space(uMAX(18 - Len("PK_" & oTableDef.name), 3)) & "PRIMARY KEY " & sClusterType & " ("
      
         'The index object can contain a collection of fields, one for each field the index contains.
         ctr = 0
         For Each fld In idx.Fields
            ctr = ctr + 1
            If ctr = idx.Fields.count Then
               DDL = DDL & fld.name & " "
            Else
               DDL = DDL & fld.name & ", "
            End If
         Next fld
         
         '''THIS NEEDS TO GO IN MAYBE
         'WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF)  ON [PRIMARY]
         
         DDL = DDL & ") , " & vbCrLf
      
      ElseIf idx.Unique = True Then         '//UNIQUE INDEX
         DDL = DDL & "  CONSTRAINT IDX_UNQ_" & oTableDef.name & Space(uMAX(18 - Len("CONSTRAINT IDX_UNQ_" & oTableDef.name), 3)) & "UNIQUE " & sClusterType & " ("
      
         'The index object can contain a collection of fields, one for each field the index contains.
         ctr = 0
         For Each fld In idx.Fields
            ctr = ctr + 1
            If ctr = idx.Fields.count Then
               DDL = DDL & fld.name & " "
            Else
               DDL = DDL & fld.name & ", "
            End If
         Next fld
         
         'slap comma out for last one
         If idxctr = oTableDef.Indexes.count Then
            DDL = DDL & ")  " & vbCrLf
         Else
            DDL = DDL & ") , " & vbCrLf
         End If
     
      
      Else                                   '//REGULAR INDEX
         DDL = DDL & "  INDEX IDX_" & idx.name & Space(uMAX(23 - Len("IDX_" & idx.name), 3)) & "" & sClusterType & " ("
      
         'The index object can contain a collection of fields, one for each field the index contains.
         For Each fld In idx.Fields
            DDL = DDL & fld.name & " "
         Next fld
         
         '''THIS NEEDS TO GO IN MAYBE
         'WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF)  ON [PRIMARY]
         
         'slap comma out for last one
         If idxctr = oTableDef.Indexes.count Then
            DDL = DDL & ")  " & vbCrLf
         Else
            DDL = DDL & ") , " & vbCrLf
         End If
         
      End If
   Next idx
   '*&*
   
   DDL = DDL & ") " & vbCrLf
   
   DDL = DDL & "ON [PRIMARY] " & vbCrLf
   DDL = DDL & "GO" & vbCrLf
   DDL = DDL & "" & vbCrLf
   
   ''Debug.Print DDL
   
   DDLer = DDL
   
End Function


Sub DeleteRelationship(tableName As String)
'//=============================================================================//
'/|        SUB:  DeleteRelationship                                             |/
'/| PARAMETERS:  TableName     Table to Remove Relationships form               |/
'/|    RETURNS:  -NONE-                                                         |/
'/|    PURPOSE:  Deletes all relationships associated with a table.             |/
'/|      USAGE:  DeleteRelationship("Employees")                                |/
'/|         BY:  Sean                                                           |/
'/|       DATE:  11/30/96                                                       |/
'/|    HISTORY:                                                                 |/
'//=============================================================================//

Dim db As Database, i As Integer
On Error GoTo DeleteRelationship_Err
  
   Set db = CurrentDb
   For i = 0 To db.Relations.count - 1
      If (db.Relations(i).Table = tableName) Or (db.Relations(i).ForeignTable = tableName) Then
         db.Relations.DELETE db.Relations(i).name
      End If
   Next i

DeleteRelationship_Done:
  Exit Sub

DeleteRelationship_Err:
  If Err.Number = 3265 Then
    MsgBox "There is no table named " & tableName & " in this database."
  Else
    MsgBox "An unexpected error (" & Err.Number & ") occurred: " & Err.Description
  End If
  Resume DeleteRelationship_Done

End Sub


Sub DeleteTable(tableName As String)
'//=============================================================================//
'/|        SUB:  DeleteTable                                                    |/
'/| PARAMETERS:  TableName     Table to Delete                                  |/
'/|    RETURNS:  -NONE-                                                         |/
'/|    PURPOSE:  Deletes a table.                                               |/
'/|      USAGE:  DeleteTable("Employees")                                       |/
'/|         BY:  Sean                                                           |/
'/|       DATE:  11/30/96                                                       |/
'/|    HISTORY:                                                                 |/
'//=============================================================================//
   
Dim db As Database, i As Integer, td As TableDef

On Error GoTo DeleteTable_Err:
   
   Set db = CurrentDb
   Set td = db.TableDefs(tableName)
   
   db.TableDefs.DELETE tableName

DeleteTable_Done:
  Exit Sub

DeleteTable_Err:
   Select Case Err
      Case 3211 'Table Currently in Use
         If MsgBox("An error (" & Err.Number & ") occurred:  " & Err.Description & " (" & tableName & ") Please hit OK to close the table continue.", MB_OKCANCEL) Then
            DoCmd.Close acTable, tableName, acSaveYes
            Resume
         Else
            Resume DeleteTable_Done
         End If
      Case Else
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume DeleteTable_Done
   End Select
End Sub


Sub DocumentDB()
'//=============================================================================//
'/|        SUB:  DocumentDB                                                     |/
'/| PARAMETERS:                                                                 |/
'/|    RETURNS:  -NONE-                                                         |/
'/|    PURPOSE:  Documents a Database's Table, Queries, Forms, etc.             |/
'/|      USAGE:  DocumentDB()                                                   |/
'/|         BY:  Sean                                                           |/
'/|       DATE:  11/30/96                                                       |/
'/|    HISTORY:                                                                 |/
'//=============================================================================//

    Dim DefaultWorkspace As Workspace
    Dim MyDatabase As Database, TempDatabase As Database
    Dim i As Integer, j As Integer
    Dim CurrentDir As String
    Set DefaultWorkspace = DBEngine.Workspaces(0)

    CurrentDir = "DBLOG.TXT"

' Open DBLOG.TXT
    Open CurrentDir For Output As #1    ' Open to write file.

' Enumerate all databases.
    For j = 0 To DefaultWorkspace.Databases.count - 1
        Set TempDatabase = DefaultWorkspace.Databases(j)
        Print #1,
        Print #1, "Enumeration of Databases("; j; "): "; TempDatabase.name
        Print #1, "   Number of TableDefs:    "; TempDatabase.TableDefs.count - 1
        Print #1, "   Number of Containers:   "; TempDatabase.Containers.count - 1
        Print #1, "   Number of QueryDefs:    "; TempDatabase.QueryDefs.count - 1
        Print #1, "   Number of Recordsets:   "; TempDatabase.Recordsets.count - 1
        Print #1, "   Number of Relations:    "; TempDatabase.Relations.count - 1
        Print #1,

' Enumerate table definitions.
        Print #1, "TableDef: Name, LastUpdated, DateCreated"
        For i = 0 To TempDatabase.TableDefs.count - 1
            Print #1, "  "; TempDatabase.TableDefs(i).name;
            Print #1, Spc(30 - Len(TempDatabase.TableDefs(i).name));
            Print #1, "  "; TempDatabase.TableDefs(i).LastUpdated;
            Print #1, Space(30 - Len(TempDatabase.TableDefs(i).LastUpdated));
            Print #1, "  "; TempDatabase.TableDefs(i).DateCreated
        Next i
        Print #1,

' Enumerate containers.
        Print #1, "Container: Name, Owner"
        For i = 0 To TempDatabase.Containers.count - 1
            Print #1, "  "; TempDatabase.Containers(i).name;
            Print #1, "  "; TempDatabase.Containers(i).Owner
        Next i
        Print #1,

' Enumerate query definitions.
        Print #1, "QueryDef: Name, LastUpdated, DateCreated"
        For i = 0 To TempDatabase.QueryDefs.count - 1
            Print #1, "  "; TempDatabase.QueryDefs(i).name;
            Print #1, Spc(30 - Len(TempDatabase.QueryDefs(i).name));
            Print #1, "  "; TempDatabase.QueryDefs(i).LastUpdated;
            Print #1, Spc(30 - Len(TempDatabase.QueryDefs(i).LastUpdated));
            Print #1, "  "; TempDatabase.QueryDefs(i).DateCreated
        Next i
        Print #1,

' Enumerate relationships.
        Print #1, "Relation: Name, Table, ForeignTable"
        For i = 0 To TempDatabase.Relations.count - 1
            Print #1, "  "; TempDatabase.Relations(i).name;
            Print #1, Space(30 - Len(TempDatabase.Relations(i).name));
            Print #1, "  "; TempDatabase.Relations(i).Table;
            Print #1, Space(30 - Len(TempDatabase.Relations(i).Table));
            Print #1, "  "; TempDatabase.Relations(i).ForeignTable
        Next i
        Print #1,
    Next j

' Close DBLOG.TXT
    Close #1                             ' Close file.

End Sub


Public Function DownloadFile(sSourceUrl As String, sLocalFile As String) As Boolean
'**********************************************
'Author  :  Ravi Poluri
'Date    :  01-21-03/v6.8.110
'Input   :
'Output  :
'Comments:
'Uses Win32 API to download a file.
'This routine eliminates the need for the Inet ActiveX control.
'**********************************************
   
   'If the API returns ERROR_SUCCESS (0),
   'return True from the function
   DownloadFile = URLDownloadToFile(0&, sSourceUrl, sLocalFile, 0&, 0&) = ERROR_SUCCESS
   
Err_Handler:
   If Err Then
'      ProcessMsg Err.Number, Err.Description, "", "DownloadFile"
      MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Function EDate(vntDate As Variant) As Variant
'//======================================================================//
'/|   FUNCTION:  EDate                                                   |/
'/| PARAMETERS:  vntDate                                                 |/
'/|    RETURNS:  3/3/94 23:59:59                                         |/
'/|    PURPOSE:  Takes a Date with or without time and returns the last  |/
'/|              of the Day for using Dates in Between statements.       |/
'/|              The Function Returns Null if an invalid Date is passed. |/
'/|      USAGE:  EDate("3/3/94") or EDate(#3/3/94#)                      |/
'/|         BY:  Sean                                                    |/
'/|       DATE:  11/30/96                                                |/
'/|    HISTORY:                                                          |/
'//======================================================================//

   If Not IsDate(vntDate) Then
      EDate = Null
   Else
      If VarType(vntDate) <> 7 Then
         vntDate = CVDate(vntDate)
      End If
      vntDate = Fix(vntDate)
      EDate = vntDate & " 23:59:59"
   End If

End Function


Public Function EncryptDecrypt(ByVal sData As String, ByVal sKey As String) As String
   'RKP/12-03-09
    Dim l As Long, i As Long, byIn() As Byte, byOut() As Byte, byKey() As Byte
    Dim bEncOrDec As Boolean
     'confirm valid string and key input:
    If Len(sData) = 0 Or Len(sKey) = 0 Then EncryptDecrypt = "Invalid argument(s) used": Exit Function
     'check whether running encryption or decryption (flagged by presence of "xxx" at start of sData):
    If Left$(sData, 3) = "xxx" Then
        bEncOrDec = False 'decryption
        sData = Mid$(sData, 4)
    Else
        bEncOrDec = True 'encryption
    End If
     'assign strings to byte arrays (unicode)
    byIn = sData
    byOut = sData
    byKey = sKey
    l = LBound(byKey)
    For i = LBound(byIn) To UBound(byIn) - 1 Step 2
        byOut(i) = ((byIn(i) + Not bEncOrDec) Xor byKey(l)) - bEncOrDec 'avoid Chr$(0) by using bEncOrDec flag
        l = l + 2
        If l > UBound(byKey) Then l = LBound(byKey) 'ensure stay within bounds of Key
    Next i
    EncryptDecrypt = byOut
    If bEncOrDec Then EncryptDecrypt = "xxx" & EncryptDecrypt 'add "xxx" onto encrypted text
End Function


Function ExcelColumnLetter2Number(ColumnLetter As String)
'?ExcelColumnLetter2Number("AZ")
'52
'?ExcelColumnLetter2Number("AA")
'27
'?ExcelColumnLetter2Number("FW")
'49  problem:  should be 153


Dim tmp As Long
Dim strLastLetter As String
Dim intColNumber As Integer
Dim intLastLetterNumber As Integer
Dim i As Integer, n As Integer

   ExcelColumnLetter2Number = 0
   
   ' Procedure: GetColumnNumber
   ' Description: This Function returns Column Number
   ' Argument: rstrL = Column Address Letter-part
   '--
   ' A-Z = 65-90 (26 Characters)
   '--
   '--
   ColumnLetter = UCase(ColumnLetter)
   
   strLastLetter = Right(ColumnLetter, 1)
   
   ' Calculate Last Letter equivalent number
   intLastLetterNumber = Asc(strLastLetter) - 64
   
   ' Find Letters-part Length
   n = Len(ColumnLetter)

   ' Calculate Column number
   intColNumber = (n - 1) * 26 + intLastLetterNumber


   ExcelColumnLetter2Number = intColNumber
    
End Function


Function ExcelColumnNumber2Letter(ColumnNumber As Long)
'?ExcelColumnNumber2Letter(26)
'Z
'?ExcelColumnNumber2Letter(52)
'AZ

Dim n As Long
Dim c As Byte
Dim s As String

   n = ColumnNumber
   Do
       c = ((n - 1) Mod 26)
       s = Chr(c + 65) & s
       n = (n - c) \ 26
   Loop While n > 0

ExcelColumnNumber2Letter = s

End Function


Sub ExecCmd(cmdline$)

    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret As Long

    'Initialize the STARTUPINFO structure:
    start.cb = Len(start)

    'Start the shelled application
    ret = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, EC_NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    'Wait for the shelled application to finish:
    ret = WaitForSingleObject(proc.hProcess, EC_INFINITE)
    ret = CloseHandle(proc.hProcess)

End Sub


Public Function ExecSQL(ByVal sql As String, ByRef ret As Long) As ADODB.Recordset
'**********************************************
'Author  :  RKP
'Date/Ver:  05-01-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    Set ExecSQL = Application.CurrentProject.Connection.Execute(sql, ret)

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function


Function ExtractSQLErr(Err_Num As Integer, Err_Str As String) As String
'//============================================================================//
'/|   FUNCTION:  ExtractSQLErr                                                 |/
'/| PARAMETERS:  Err_Num, Err_Str                                              |/
'/|    RETURNS:  -NONE-                                                        |/
'/|    PURPOSE:  This routine parses the error string returned from ODBC and   |/
'/|              extracts only the error value assigned by RAISERROR() or      |/
'/|              MSSQL-Server                                                  |/
'/|      USAGE:  ExtractSQLErr(Err_Num, Err_Str)                               |/
'/|         BY:  Sean                                                          |/
'/|       DATE:  11/30/96                                                      |/
'/|    HISTORY:                                                                |/
'//============================================================================//

Dim SQL_Err As String, Xstart  As Integer, Xlen As Integer
    
   Xstart = InStr(1, Err_Str, "#") + 1                     '// extract error value from string
   Xlen = InStr(Xstart, Err_Str, ")") - Xstart
   SQL_Err = Mid(Err_Str, Xstart, Xlen)
   
   Xstart = InStr(1, Err_Str, "[SQL Server]") + 13         '// extract error message from string
   Xlen = InStr(Xstart, Err_Str, "(#") - Xstart
   SQL_Err = SQL_Err & "; " & Mid(Err_Str, Xstart, Xlen)
   
   ExtractSQLErr = SQL_Err
   
End Function


Function FieldCruise()
    
    Dim DefaultWorkspace As Workspace
    Dim MyDatabase As Database
    Dim MyTableDef As TableDef
    Dim MyQueryDef As QueryDef
    Dim MyField As Field
    Dim i As Integer
    Dim j As Integer

    Set DefaultWorkspace = DBEngine.Workspaces(0)
    Set MyDatabase = DefaultWorkspace.Databases(0)
    
    'TABLE LOOP
    For i = 0 To MyDatabase.TableDefs.count - 1
        Set MyTableDef = MyDatabase.TableDefs(i)
        If InStr(MyTableDef.name, "tblintfintChangesASFControlVO") > 0 Then
            Debug.Print MyTableDef.name
            'FIELDS LOOP
            For j = 0 To MyTableDef.Fields.count - 1
                'Set MyField = MyTableDef.Fields(j)
               ' If MyField.Name = "SGRADE_CODE" Then
                   
                   Debug.Print j & " " & MyTableDef.Fields(j).name
               ' End If
            Next j
        End If
    Next i
End Function


Function FieldExists(strFieldName As String, strTableName As String) As Boolean
' Comments   : Returns True if field strFieldName Exists in table strTableName , False if it does not
' Parameters : FieldExists - Name of the file (use a fully qualified path)
' Returns    : Boolean
' Sample Call: ? FieldExists("Name_Address","ADDRESS_3")
' Created    : 2007/VanDamme Associates, Inc (www.vandamme.com)
' --------------------------------------------------
'?FieldExists("qrepMainWithPrc","Revenue")
On Error GoTo FieldExists_Err
Dim strTemp As String
   FieldExists = False
   strTemp = ""
   FieldExists = False
   
   
   If IsNull(CurrentDb.TableDefs(strFieldName).Fields(strTableName).name) Then
   Else
      strTemp = CurrentDb.TableDefs(strFieldName).Fields(strTableName).name
   End If

   If IsNull(CurrentDb.QueryDefs(strFieldName).Fields(strTableName).name) Then
   Else
      strTemp = CurrentDb.QueryDefs(strFieldName).Fields(strTableName).name
   End If
   
   If Len(strTemp) >= 1 Then
      FieldExists = True
   End If
      
Exit Function

FieldExists_Err:
   Resume Next

End Function


Function FileCopy(SourceName As String, DestName As String) As Integer
'
' Copies a single file SourceName to DestName
'
' Calling convention:
'   X = FileCopy("C:\This.Exe", "C:\That.Exe")
'   X = FileCopy("C:\This.Exe", "C:\Temp\This.Exe")
'   X = FileCopy("C:\OPTMODELS\PC30\CAPDATA 2006.04.11.mdb", "C:\OPTMODELS\PC30\CAPDATA.MDB")


Dim RetVal As Long

RetVal = CopyFileX(SourceName, DestName) ', 1)

If RetVal = 0 Then ' failure
    MsgBox "Copy failed -- C:\DestinationFile.txt already exists."
Else ' success
    MsgBox "Copy succeeded."
End If

End Function















Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    
      ''      Look for a file named MyFile.mdb in the Data folder:
      ''          FileExists ("C:\Data\MyFile.mdb")
      ''      Look for a folder named System in the Windows folder on C: drive:
      ''          FolderExists ("C:\Windows\System")
      ''      Look for a file named MyFile.txt on a network server:
      ''          FileExists ("\\MyServer\MyPath\MyFile.txt")
      ''      Check for a file or folder name Wotsit on the server:
      ''          FileExists("\\MyServer\Wotsit", True)
      ''      Check the folder of the current database for a file named GetThis.xls:
      ''          FileExists (TrailingSlash(CurrentProject.Path) & "GetThis.xls")
    
    ''' ?FileExists("C:\BIN\Autoexec.bat1")
    ''' ?FileExists("C:\BIN\Autoexec.bat")

    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function


Private Function FillDir(colDirList As Collection, ByVal strFolder As String, strFileSpec As String, _
    bIncludeSubfolders As Boolean)
    'Build up a list of files, and then add add to this list, any additional folders
    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant
    'Add the files to the folder.
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & strFileSpec)
    Do While strTemp <> vbNullString
        colDirList.Add strFolder & strTemp
        strTemp = Dir
    Loop
    If bIncludeSubfolders Then
        'Build collection of additional subfolders.
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0& Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop
        'Call function recursively for each subfolder.
        For Each vFolderName In colFolders
            Call FillDir(colDirList, strFolder & TrailingSlash(vFolderName), strFileSpec, True)
        Next vFolderName
    End If
End Function


Public Function FixColumnWidthsOfQuery(stName As String)
    Dim db As Database
    Dim qdf As QueryDef
    Dim fld As dao.Field
    Dim frm As Form
    Dim ictl As Integer
    Dim ctl As Control
    
    Set db = CurrentDb
    Set qdf = db.QueryDefs(stName)
    DoCmd.OpenQuery stName, acViewNormal
    Set frm = Screen.ActiveDatasheet
    For ictl = 0 To frm.Controls.count - 1
     Set ctl = frm.Controls(ictl)
     ctl.ColumnWidth = -2
     Call SetDAOFieldProperty(qdf.Fields(ictl), _
      "ColumnWidth", ctl.ColumnWidth, dbInteger)
    Next ictl
    DoCmd.Save acQuery, stName
End Function


Public Function FixColumnWidthsOfTable(stName As String)
    Dim db As Database
    Dim tdf As TableDef
    Dim fld As dao.Field
    Dim frm As Form
    Dim ictl As Integer
    Dim ctl As Control
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(stName)
    DoCmd.OpenTable stName, acViewNormal
    Set frm = Screen.ActiveDatasheet
    For ictl = 0 To frm.Controls.count - 1
     Set ctl = frm.Controls(ictl)
     ctl.ColumnWidth = -2
     Call SetDAOFieldProperty(tdf.Fields(ictl), _
      "ColumnWidth", ctl.ColumnWidth, dbInteger)
    Next ictl
'    DoCmd.Save acTable, stName
    DoCmd.Close acTable, stName, acSaveYes
End Function


Function FolderExists(strPath As String) As Boolean
    ''' ?FolderExists("C:\BIN")
    ''' ?FolderExists("C:\BING")

    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function


Function FolderCruise()

'//  example of how to loop through the files in a directory

'//  NOTE: Examples that follow demonstrate the use of this function in a Visual Basic for Applications (VBA) module.
'//  For more information about working with VBA, select Developer Reference in the drop-down list next to Search and
'//  enter one or more terms in the search box.

'//  This example uses the Dir function to check if certain files and directories exist. On the Macintosh, â€œHD:â€ is
'//  the default drive name and portions of the pathname are separated by colons instead of backslashes. Also, the
'//  Windows wildcard characters are treated as valid file-name characters on the Macintosh. However, you can use the
'//  MacID function to specify file groups.

   Dim MyFile As String, MyPath As String, MyName As String
   
   ' Returns "WIN.INI" (on Microsoft Windows) if it exists.
   MyFile = Dir("C:\WINDOWS\WIN.INI")
   
   ' Returns filename with specified extension. If more than one *.ini
   ' file exists, the first file found is returned.
   MyFile = Dir("C:\WINDOWS\*.INI")
   
   ' Call Dir again without arguments to return the next *.INI file in the
   ' same directory.
   MyFile = Dir
   
   ' Return first *.TXT file with a set hidden attribute.
   MyFile = Dir("*.TXT", vbHidden)
   
   ''MyPath = "C:\SOFTWARE\"
   ''MyPath = "C:\REMOVAL\"
   
   ''' Display the names in C:\ that represent directories.
   ''MyPath = "C:\"    ' Set the path.
   ''MyName = Dir(MyPath, vbDirectory)    ' Retrieve the first entry.
   
   MyPath = "C:\Users\smacder.NAIPAPER\Desktop\Text File Proc\"
   MyName = Dir(MyPath)    ' Retrieve the first entry.
   
   Do While MyName <> ""   ' Start the loop.
      If Left(MyName, 13) <> "ATextFileProc" And InStr(1, MyName, "PIPE") >= 1 Then
         Debug.Print MyName   ' Display entry only if it
      End If
      MyName = Dir         ' Get next entry.
   Loop
   
'''''''   Do While MyName <> ""    ' Start the loop.
'''''''      ' Ignore the current directory and the encompassing directory.
'''''''      If MyName <> "." And MyName <> ".." Then
'''''''         ' Use bitwise comparison to make sure MyName is a directory.
'''''''         If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
'''''''            Debug.Print MyName    ' Display entry only if it
'''''''         End If    ' it represents a directory.
'''''''      End If
'''''''      MyName = Dir    ' Get next entry.
'''''''   Loop


End Function


Sub FolderCruise2()
   Dim strFileName As String
   Dim intNumberOfFiles As Integer

   intNumberOfFiles = 0

   strFileName = Dir("C:\TempRich1\*  ", vbNormal)
   strFileName = Dir("C:\SOFTWARE\*.* ", vbNormal)
   strFileName = Dir("C:\REMOVAL\*.*  ", vbNormal)
   strFileName = Dir("C:\Users\smacder.NAIPAPER\Desktop\Text File Proc\*.* ", vbNormal)
   
   Do Until strFileName = ""
      intNumberOfFiles = intNumberOfFiles + 1
      If Left(strFileName, 13) <> "ATextFileProc" And InStr(1, strFileName, "PIPE") >= 1 Then
         Debug.Print strFileName   ' Display entry only if it
      End If
      strFileName = Dir()
   Loop

   MsgBox ("You Looped Through " & intNumberOfFiles & " Files")
End Sub


Public Function FormatDate(ByVal inputDate As Date, ByVal id As Integer) As String
'**********************************************
'Author  :  RKP
'Date/Ver:  05-01-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    Select Case id
        Case 1
            FormatDate = VBA.Year(inputDate) & VBA.Format(VBA.Month(inputDate), "00")
    End Select

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function


Public Function FormatFilePath(ByVal FilePath As String) As String
'**********************************************
'Author  :  RKP
'Date/Ver:  05-01-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    FormatFilePath = VBA.IIf(VBA.Right(VBA.Trim(FilePath), 1) = "\", VBA.Trim(FilePath), VBA.Trim(FilePath) & "\")

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function


Function FormatSQL(sql As String)
'Formats a SQL string flush left
   'remove any white space beyond space(1)
   Dim i As Long
   For i = 40 To 2 Step -1
    sql = Replace(sql, Space(i + 1), " ")
   Next
   'format sql flush left
   sql = Replace(sql, "FROM", vbNewLine & "FROM")
   sql = Replace(sql, "WHERE", vbNewLine & "WHERE")
   sql = Replace(sql, "AND", vbNewLine & "AND" & vbNewLine + vbTab)
   sql = Replace(sql, "HAVING", vbNewLine & "HAVING")
   sql = Replace(sql, "GROUP BY", vbNewLine & "GROUP BY")
   sql = Replace(sql, "ORDER BY", vbNewLine & "ORDER BY")
   sql = Replace(sql, "OR ", vbNewLine & "OR" & vbNewLine & vbTab)
   sql = Replace(sql, "INNER JOIN", vbNewLine & "INNER JOIN")
   sql = Replace(sql, "SET ", vbNewLine & "SET")
   sql = Replace(sql, "VALUES", vbNewLine & "VALUES")
   FormatSQL = sql
End Function


Public Function FormatStr(ByVal strText As String, ParamArray Args()) As String
  ' Comments:  Format a string like .NET's String.Format
  '
  ' Params  :  Format string with place holder values like {0} and {1}
  ' Returns :  Formated string

' © codekabinett.com - You may use, modify, copy, distribute this code as long as this line remains
' https://codekabinett.com/index.php?Lang=2
' https://codekabinett.com/rdumps.php?Lang=2&targetDoc=vba-printf-string-format-function


    Dim i           As Integer
    Dim strRetVal   As String
    Dim startPos    As Integer
    Dim endPos      As Integer
    Dim formatString As String
    Dim argValueLen As Integer
    strRetVal = strText
    
    For i = LBound(Args) To UBound(Args)
        argValueLen = Len(CStr(i))
        startPos = InStr(strRetVal, "{" & CStr(i) & ":")
        If startPos > 0 Then
            endPos = InStr(startPos + 1, strRetVal, "}")
            formatString = Mid(strRetVal, startPos + 2 + argValueLen, endPos - (startPos + 2 + argValueLen))
            strRetVal = Mid(strRetVal, 1, startPos - 1) & Format(Nz(Args(i), ""), formatString) & Mid(strRetVal, endPos + 1)
        Else
            strRetVal = Replace(strRetVal, "{" & CStr(i) & "}", Nz(Args(i), ""))
        End If
    Next i

    FormatStr = strRetVal

End Function


Function GetAccessEXEVersion() As String
'Valid for us with Access 2000 or later.
'Original version may be from Tom van Stiphout, not sure?

'SysCmd(715) -> 6606
'Application.Version OR SysCmd(acSysCmdAccessVer) -> 12.0
    On Error Resume Next
    Dim sAccessVerNo As String
'    sAccessVerNo = fGetProductVersion(Application.SysCmd(acSysCmdAccessDir) & "msaccess.exe")
    sAccessVerNo = SysCmd(acSysCmdAccessVer) & "." & SysCmd(715)
    Select Case sAccessVerNo
        'Access 2000
        Case "9.0.0.0000" To "9.0.0.2999": GetAccessEXEVersion = "Microsoft Access 2000 - Build:" & sAccessVerNo
        Case "9.0.0.3000" To "9.0.0.3999": GetAccessEXEVersion = "Microsoft Access 2000 SP1 - Build:" & sAccessVerNo
        Case "9.0.0.4000" To "9.0.0.4999": GetAccessEXEVersion = "Microsoft Access 2000 SP2 - Build:" & sAccessVerNo
        Case "9.0.0.6000" To "9.0.0.6999": GetAccessEXEVersion = "Microsoft Access 2000 SP3 - Build:" & sAccessVerNo
        'Access 2002
        Case "10.0.2000.0" To "10.0.2999.9": GetAccessEXEVersion = "Microsoft Access 2002 - Build:" & sAccessVerNo
        Case "10.0.3000.0" To "10.0.3999.9": GetAccessEXEVersion = "Microsoft Access 2002 SP1 - Build:" & sAccessVerNo
        Case "10.0.4000.0" To "10.0.4999.9": GetAccessEXEVersion = "Microsoft Access 2002 SP2 - Build:" & sAccessVerNo
        'Access 2003
        Case "11.0.0000.0" To "11.0.5999.9999": GetAccessEXEVersion = "Microsoft Access 2003 - Build:" & sAccessVerNo
        Case "11.0.6000.0" To "11.0.6999.9999": GetAccessEXEVersion = "Microsoft Access 2003 SP1 - Build:" & sAccessVerNo
        Case "11.0.7000.0" To "11.0.7999.9999": GetAccessEXEVersion = "Microsoft Access 2003 SP2 - Build:" & sAccessVerNo
        Case "11.0.8000.0" To "11.0.8999.9999": GetAccessEXEVersion = "Microsoft Access 2003 SP3 - Build:" & sAccessVerNo
        'Access 2007
        Case "12.0.0000.0" To "12.0.5999.9999": GetAccessEXEVersion = "Microsoft Access 2007 - Build:" & sAccessVerNo
        Case "12.0.6000.0" To "12.0.6422.9999": GetAccessEXEVersion = "Microsoft Access 2007 SP1 - Build:" & sAccessVerNo
        Case "12.0.6423.0" To "12.0.5999.9999": GetAccessEXEVersion = "Microsoft Access 2007 SP2 - Build:" & sAccessVerNo
        'Unable to locate specific build versioning for SP3 - to be validated at a later date.
        '  Hopefully MS will eventually post the info on their website?!
        Case "12.0.6000.0" To "12.0.9999.9999": GetAccessEXEVersion = "Microsoft Access 2007 SP3 - Build:" & sAccessVerNo
        'Access 2010
        Case "14.0.0000.0000" To "14.0.6022.1000": GetAccessEXEVersion = "Microsoft Access 2010 - Build:" & sAccessVerNo
        Case "14.0.6023.1000" To "14.0.7014.9999": GetAccessEXEVersion = "Microsoft Access 2010 SP1 - Build:" & sAccessVerNo
        Case "14.0.7015.1000" To "14.0.9999.9999": GetAccessEXEVersion = "Microsoft Access 2010 SP2 - Build:" & sAccessVerNo
        'Access 2013
        Case "15.0.0000.0000" To "15.0.4569.1505": GetAccessEXEVersion = "Microsoft Access 2013 - Build:" & sAccessVerNo
        Case "15.0.4569.1506" To "15.0.9999.9999": GetAccessEXEVersion = "Microsoft Access 2013 SP1 - Build:" & sAccessVerNo
        Case "16.0.0000.0000" To "16.0.4312.0999": GetAccessEXEVersion = "Microsoft Access 2016 - Build:" & sAccessVerNo
        Case "16.0.4312.1000" To "16.0.9999.9999": GetAccessEXEVersion = "Microsoft Access 2016 Update - Build:" & sAccessVerNo
        Case Else: GetAccessEXEVersion = "Unknown Version"
    End Select
    If SysCmd(acSysCmdRuntime) Then GetAccessEXEVersion = GetAccessEXEVersion & " Run-time"
End Function


Function GetAccessFileVersion() As String
   Dim intFormat As Integer
   Dim s As String
   intFormat = CurrentProject.FileFormat
   GetAccessFileVersion = False
   
   Select Case intFormat
       Case 2
         s = "Microsoft Access 2"
       Case 7
         s = "Microsoft Access 95"
       Case 8
         s = "Microsoft Access 97"
       Case 9
         s = "Microsoft Access 2000"
       Case 10
         s = "Microsoft Access 2002"
       Case 11
         s = "Microsoft Access 2003"
       Case 12
         s = "Microsoft Access 2007"
       Case 14
         s = "Microsoft Access 2010"
       Case 15
         s = "Microsoft Access 2013"
       Case 16
         s = "Microsoft Access 2016"
       Case Else
         s = "Unknown or Higher Than Microsoft Access 2016"
   End Select
   GetAccessFileVersion = s
End Function


Sub GetAccessVersionTest()
   MsgBox "You are currently running " & Application.name _
      & " EXE version " & Application.Version & ", build " _
      & Application.Build & "." & vbCrLf & vbCrLf _
      & "You are currently in an Access file of version " _
      & GetAccessFileVersion & "."

End Sub


Function GetAccessHwnd() As Long   'Get the hWnd of the Main Access Window
     
Dim lnghWnd As Long
Dim lngAccesshWnd As Long

   lnghWnd = GetActiveWindow()       '// Get hWnd of active window
   lngAccesshWnd = lnghWnd

   While lnghWnd <> 0
      lngAccesshWnd = lnghWnd        '// Keep getting parents until parent = 0
      lnghWnd = GetParent(lnghWnd)
   Wend

   GetAccessHwnd = lngAccesshWnd
        
End Function


Function GetCOPTpath() As String
'//==============================================================//
'/|   FUNCTION:  GetCOPTpath                                     |/
'/| PARAMETERS:  -NONE-                                          |/
'/|    RETURNS:  String, the name of path                        |/
'/|    PURPOSE:  Return the name of the path to C-OPT location   |/
'/|      USAGE:  s = GetCOPTpath()                               |/
'/|         BY:  Sean                                            |/
'/|       DATE:  11/17/16                                        |/
'/|    HISTORY:                                                  |/
'//==============================================================//

Dim i           As Integer
Dim sFullPath   As String
Dim strFile     As String
Dim strCOPTfile As String

'' USAGE:
'' strFile = ListFiles("C:\Program Files\BMOS\C-OPT\", "C-OPTConsole.exe", True)
'' strFile = ListFiles("C:\Program Files\",            "C-OPTConsole.exe", True)
'' strFile = ListFiles("C:\Program Files (x86)\",      "C-OPTConsole.exe", True)
'' strFile = ListFiles("C:\Program Files (x86)\",      "C-OPTConsole.exe", True)
 
'' Or if you are searching a network shared drive, you can do it this way:
'' strFile = ListFiles(\\servername\folder\, "example.doc", True)
 
'' IMPORTANT:   this code may be slow to run if it has to search thousands of files.
'' If possible, make your starting location as close as possible to where you think
'' the file will be.

   GetCOPTpath = ""
   
   '64 bit
   strFile = ""
   If Len(strFile) <= 0 Then strFile = ListFiles("C:\Programs\MATH\C-OPTx64E\", "C-OPTConsole.EXE", True)
   If Len(strFile) <= 0 Then strFile = ListFiles("C:\Program Files\BMOS\C-OPT\", "C-OPTConsole.EXE", True)
   If Len(strFile) <= 0 Then strFile = ListFiles("C:\Program Files\BMOS\C-OPT (64-bit)\", "C-OPTConsole.EXE", True)
   If Len(strFile) <= 0 Then strFile = ListFiles("C:\Program Files (x86)\BMOS\C-OPT\", "C-OPTConsole.EXE", True)
   If Len(strFile) <= 0 Then strFile = ListFiles("C:\Programs\MATH\C-OPTx64\", "C-OPTConsole.EXE", True)
                                                 'C:\Programs\MATH\C-OPTx64E
   
   If Len(strFile) > 0 Then
      sFullPath = strFile
   End If

   For i = Len(sFullPath) To 1 Step -1               '// Search backwards in string for backslash character //
      If Mid(sFullPath, i, 1) = "\" Then             '// Did we find a backslash? //
         sFullPath = Left(sFullPath, i)              '// then return everything to left of it. //
         GetCOPTpath = sFullPath
         Exit Function
      End If
   Next i

End Function


Function GetCurrentMDBName() As String
'//=====================================================================//
'/|   FUNCTION:  GetCurrentMDBName                                      |/
'/| PARAMETERS:  -NONE-                                                 |/
'/|    RETURNS:  String, the name of the current .MDB i.e. STM_LIB.mdb  |/
'/|    PURPOSE:  Return the name of the Current .MDB without the Path   |/
'/|      USAGE:  s = GetCurrentMDBName()                                |/
'/|         BY:  Sean                                                   |/
'/|       DATE:  11/30/96                                               |/
'/|    HISTORY:                                                         |/
'//=====================================================================//

Dim i As Integer, FullPath As String
   
   FullPath = CurrentDb.name
   
   For i = Len(FullPath) To 1 Step -1               '// Search backwards in string for backslash character //
      If Mid(FullPath, i, 1) = "\" Then             '// Did we find a backslash? //
         GetCurrentMDBName = Mid(FullPath, i + 1)   '// then return everything to right of it! //
         Exit Function
      End If
   Next i

End Function


Function GetCurrentMDBwPath()
'//=========================================================================//
'/|   FUNCTION:  GetCurrentMDBwPath                                         |/
'/| PARAMETERS:  -NONE-                                                     |/
'/|    RETURNS:  String, the current .MDB i.e. C:\My Documents\STM_LIB.mdb  |/
'/|    PURPOSE:  Return the name of the Current .MDB with the Path          |/
'/|      USAGE:  s = GetCurrentMDBwPath()                                   |/
'/|         BY:  Sean                                                       |/
'/|       DATE:  11/30/96                                                   |/
'/|    HISTORY:                                                             |/
'//=========================================================================//

Dim MyDB As Database

   Set MyDB = CurrentDb()               '// NOTE:  There's no final backslash
   GetCurrentMDBwPath = MyDB.name       '//        returned.

End Function


Function GetFileName(strFullPathAndName As String) As String
'//=====================================================================//
'/|   FUNCTION:  GetFileName                                            |/
'/| PARAMETERS:  strFullPathAndName, string, a fully qualified path     |/
'/|                 and file name                                       |/
'/|    RETURNS:  String, the name of the current file i.e. SALES.XLS    |/
'/|    PURPOSE:  Return the name of the current file without the Path   |/
'/|      USAGE:  s = GetFileName("C:\DDI\PLANNING\SALES.XLS")           |/
'/|         BY:  Sean                                                   |/
'/|       DATE:  1/11/97                                                |/
'/|    HISTORY:                                                         |/
'//=====================================================================//
   
   Dim i As Integer, FullPath As String
   FullPath = strFullPathAndName
   
   For i = Len(FullPath) To 1 Step -1               '// Search backwards in string for backslash character //
      If Mid(FullPath, i, 1) = "\" Then             '// Did we find a backslash? //
         GetFileName = Mid(FullPath, i + 1)         '// then return everything to right of it! //
         Exit Function
      End If
   Next i

End Function


Public Function GetFileSize(ByVal FilePath As String, ByVal fileSize As fileSizeEnum, Optional numFormat = "0.00") As Double
'**********************************************
'Author  :  RKP
'Date/Ver:  01-06-11/V01
'Input   :
'Output  :
'Comments:
'**********************************************
   On Error GoTo Err_Handler
 
   Select Case fileSize
      Case fileSizeEnum.Bytes
         GetFileSize = VBA.Format(CreateObject("Scripting.FileSystemObject").GetFile(FilePath).Size, numFormat)
      Case fileSizeEnum.KB
         GetFileSize = VBA.Format(CreateObject("Scripting.FileSystemObject").GetFile(FilePath).Size / 1024, numFormat)
      Case fileSizeEnum.MB
         GetFileSize = VBA.Format((CreateObject("Scripting.FileSystemObject").GetFile(FilePath).Size / 1024) / 1024, numFormat)
      Case fileSizeEnum.GB
         GetFileSize = VBA.Format(((CreateObject("Scripting.FileSystemObject").GetFile(FilePath).Size / 1024) / 1024) / 1024, numFormat)
      Case fileSizeEnum.TB
         GetFileSize = VBA.Format((((CreateObject("Scripting.FileSystemObject").GetFile(FilePath).Size / 1024) / 1024) / 1024) / 1024, numFormat)
      Case Else
         GetFileSize = VBA.Format(CreateObject("Scripting.FileSystemObject").GetFile(FilePath).Size, numFormat)
   End Select
   

Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", ""
      MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Public Function GetFileVersion(ByVal FilePath As String) As String
'**********************************************
'Author  :  RKP
'Date/Ver:  01-07-2010/V01
'Input   :
'Output  :
'Comments:
'GetFileVersion(VBA.Environ("SYSTEMROOT") & "\system32\msi.dll")
'**********************************************
   On Error GoTo Err_Handler

   GetFileVersion = CreateObject("Scripting.FileSystemObject").GetFileVersion(FilePath)

Err_Handler:
   mlLastErr = Err.Number
   msLastErr = Err.Description
   'Function1 = mlLastErr
   If Err Then
      If Err.Number = 49 Then 'Bad DLL calling convention
         mlLastErr = 0
         msLastErr = ""
         Resume Next
      Else
         'ProcessMsg Err.Number, Err.Description, "", ""
         MsgBox Err.Number & " - " & Err.Description
      End If
   End If
   Exit Function
   Resume
End Function


Function GetPathName(strFullPathAndName As String) As String
'//===========================================================================//
'/|   FUNCTION:  GetPathName                                                  |/
'/| PARAMETERS:  strFullPathAndName, string, a fully qualified path           |/
'/|                 and file name                                             |/
'/|    RETURNS:  String, the name of the current path i.e. C:\DDI\PLANNING\   |/
'/|    PURPOSE:  Return the name of the current file without the Path         |/
'/|      USAGE:  s = GetPathName("C:\DDI\PLANNING\SALES.XLS")                 |/
'/|         BY:  Sean                                                         |/
'/|       DATE:  1/11/97                                                      |/
'/|    HISTORY:                                                               |/
'//===========================================================================//
   
   Dim i As Integer, FullPath As String
   FullPath = strFullPathAndName
   
   For i = Len(FullPath) To 1 Step -1               '// Search backwards in string for backslash character //
      If Mid(FullPath, i, 1) = "\" Then             '// Did we find a backslash? //
         GetPathName = Left(FullPath, i)            '// then return IT AND everything to left of it! //
         Exit Function
      End If
   Next i

End Function


Function GetFileDateTime(ByVal FileName As String) As Variant
'//=====================================================//
'/|   FUNCTION: GetFileDateTime                         |/
'/| PARAMETERS: Filename                                |/
'/|    RETURNS: Date and Time                           |/
'/|    PURPOSE: Determine DateTimeStamp of File         |/
'/|      USAGE: d = GetFileDateTime("C:\Config.sys")    |/
'/|         BY: Sean                                    |/
'/|       DATE: 1/10/97                                 |/
'/|    HISTORY:                                         |/
'//=====================================================//
   
   Dim ofs As typOFSTRUCT
   Dim iDate As Long
   Dim iTime As Long

   Const DAY_MASK = &H1F
   Const MONTH_MASK = &H1E0
   Const YEAR_MASK = &HFE00

   Const SECOND_MASK = &H1F
   Const MINUTE_MASK = &H7E0
   Const HOUR_MASK = &HF800

   If WinOpenFile(FileName, ofs, OF_EXIST) <> -1 Then
      iDate = Asc(Mid$(ofs.szReserved, 2, 1)) * 256& + Asc(Mid$(ofs.szReserved, 1, 1))
      iTime = Asc(Mid$(ofs.szReserved, 4, 1)) * 256& + Asc(Mid$(ofs.szReserved, 3, 1))
      GetFileDateTime = DateSerial(((iDate And YEAR_MASK) \ &H200) + 1980, (iDate And MONTH_MASK) \ &H20, (iDate And DAY_MASK)) + TimeSerial((iTime And HOUR_MASK) \ &H800, (iTime And MINUTE_MASK) \ &H20, (iTime And SECOND_MASK) * 2)
   Else
      GetFileDateTime = Null
   End If

End Function


Function GetHaversineMiles(lat1Degrees As Double, lon1Degrees As Double, lat2Degrees As Double, lon2Degrees As Double) As Double
    Dim earthSphereRadiusKilometers As Double
    Dim kilometerConversionToMilesFactor As Double
    Dim lat1Radians As Double
    Dim lon1Radians As Double
    Dim lat2Radians As Double
    Dim lon2Radians As Double
    Dim AsinBase As Double
    Dim DerivedAsin As Double
    'Mean radius of the earth (replace with 3443.89849 to get nautical miles)
    earthSphereRadiusKilometers = 6371
    'Convert kilometers into miles
    kilometerConversionToMilesFactor = 0.621371
    'Convert each decimal degree to radians
    lat1Radians = (lat1Degrees / 180) * 3.14159265359
    lon1Radians = (lon1Degrees / 180) * 3.14159265359
    lat2Radians = (lat2Degrees / 180) * 3.14159265359
    lon2Radians = (lon2Degrees / 180) * 3.14159265359
    AsinBase = Sin(Sqr(Sin((lat1Radians - lat2Radians) / 2) ^ 2 + Cos(lat1Radians) * Cos(lat2Radians) * Sin((lon1Radians - lon2Radians) / 2) ^ 2))
    DerivedAsin = (AsinBase / Sqr(-AsinBase * AsinBase + 1))
    'Get distance from [lat1,lon1] to [lat2,lon2]
    'KM:    = Round(2 * DerivedAsin * earthSphereRadiusKilometers, 2)
    'Miles: = Round(2 * DerivedAsin * (earthSphereRadiusKilometers * kilometerConversionToMilesFactor), 2)
    GetHaversineMiles = Round(2 * DerivedAsin * (earthSphereRadiusKilometers * kilometerConversionToMilesFactor), 2)
End Function


Function GetOpenFileDialog() As String
'//=============================================================//
'/|   FUNCTION: GetOpenFileDialog                               |/
'/| PARAMETERS: -NONE-                                          |/
'/|    RETURNS: "C:\Program Files\Common Files\Msgraph.hlp"     |/
'/|    PURPOSE: Get Open File Name from User                    |/
'/|      USAGE: strFileToOpen = GetOpenFileDialog()             |/
'/|         BY: Sean                                            |/
'/|       DATE: 11/30/96                                        |/
'/|    HISTORY:                                                 |/
'//=============================================================//

Dim strMessage, strFilter, strFileName, strFileTitle, strDefExt, strTitle, strCurDir As String
Dim intAPIResults As Integer

    GetOpenFileDialog = False
    '// Define the filter string and allocate space in the "c" string //
    strFilter = "Porter(*.csv)" & Chr$(0) & "*.CSV" & Chr$(0)
    strFilter = strFilter & "Text(*.txt)" & Chr$(0) & "*.TXT" & Chr$(0)
    
    'strFilter = strFilter & "Access(*.mdb)" & Chr$(0) & "*.MDB;*.MDA" & Chr$(0)
    'strFilter = strFilter & "Batch(*.bat)" & Chr$(0) & "*.BAT" & Chr$(0)
    strFilter = strFilter & Chr$(0)
    
    strFileName = Chr$(0) & Space$(255) & Chr$(0)             '// Allocate Space for String //
    strFileTitle = Space$(255) & Chr$(0)                      '// Allocate Space for String //
    strTitle = "Open File" & Chr$(0)                          '// Give the dialog a caption title. //
    strDefExt = "TXT" & Chr$(0)                               '// Default Extension = TXT //
    strCurDir = CurDir$ & Chr$(0)                             '// Set up the defualt directory //
    
    OPENFILENAME.lStructSize = Len(OPENFILENAME)
    OPENFILENAME.hWndOwner = 0&                    '// If called from Form use = Screen.ActiveForm.hWnd //
'    OPENFILENAME.hwndOwner = Screen.ActiveForm.Hwnd
    OPENFILENAME.lpstrFilter = strFilter
    OPENFILENAME.nFilterIndex = 1
    OPENFILENAME.lpstrFile = strFileName
    OPENFILENAME.nMaxFile = Len(strFileName)
    OPENFILENAME.lpstrFileTitle = strFileTitle
    OPENFILENAME.nMaxFileTitle = Len(strFileTitle)
    OPENFILENAME.lpstrTitle = strTitle
    OPENFILENAME.flags = OFN_FILEMUSTEXIST Or OFN_ALLOWMULTISELECT
    OPENFILENAME.lpstrDefExt = strDefExt
    OPENFILENAME.hInstance = 0
    OPENFILENAME.lpstrCustomFilter = Chr$(0)
    OPENFILENAME.nMaxCustFilter = 0
    OPENFILENAME.lpstrInitialDir = strCurDir
    OPENFILENAME.nFileOffset = 0
    OPENFILENAME.nFileExtension = 0
    OPENFILENAME.lCustData = 0
    OPENFILENAME.lpfnHook = 0
    OPENFILENAME.lpTemplateName = Chr$(0)

    intAPIResults = GetOpenFileName(OPENFILENAME)

    If intAPIResults <> 0 Then
       strFileName = Left$(OPENFILENAME.lpstrFile, InStr(OPENFILENAME.lpstrFile, Chr$(0)) - 1)
       strMessage = "The file you chose was " & strFileName
    Else
       strMessage = "No file was selected"
    End If

    '// MsgBox strMessage//
    
    GetOpenFileDialog = strFileName
End Function


Function GetScreenResolution() As String
'//=====================================================//
'/|   FUNCTION: GetScreenResolution                     |/
'/| PARAMETERS: -NONE-                                  |/
'/|    RETURNS: i.e. 640x480, 800x600, 1024x768         |/
'/|    PURPOSE: Determine Current Screen Size           |/
'/|      USAGE: l = GetScreenResolution()               |/
'/|         BY: Sean                                    |/
'/|       DATE: 11/30/96                                |/
'/|    HISTORY:                                         |/
'//=====================================================//
    
Dim R As typRECT
Dim hWnd As Long
Dim RetVal As Long
    
    hWnd = GetDesktopWindow()
    RetVal = GetWindowRect(hWnd, R)
    GetScreenResolution = (R.x2 - R.x1) & "x" & (R.y2 - R.y1)

End Function


Public Function GetSetting(ByVal inputValue As String, Optional tableName = "tbl000Settings", Optional inputFieldName = "KeyName", Optional outputFieldName = "KeyValue")
'**********************************************
'Author  :  RKP
'Date/Ver:  03-18-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    GetSetting = Application.CurrentProject.Connection.Execute("SELECT [" & outputFieldName & "] FROM [" & tableName & "] WHERE [" & inputFieldName & "] = '" & inputValue & "'").Fields(0).value & ""

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            'MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function


Function getSQL(strQueryName As String) As Variant
'//==============================================================//
'/|   FUNCTION:  GetSQL                                          |/
'/| PARAMETERS:  strQueryName, the Name of the Query to Check    |/
'/|    RETURNS:  String of SQL in the Query, Null On Failure     |/
'/|    PURPOSE:  Returns the SQL String of a Query               |/
'/|      USAGE:  GetSQL("Query1")                                |/
'/|         BY:  Sean                                            |/
'/|       DATE:  11/30/96                                        |/
'/|    HISTORY:                                                  |/
'//==============================================================//

Dim db As Database, qd As QueryDef
On Error GoTo GetSQL_Err

    getSQL = Null                           '// Default - Failed
    Set db = DBEngine(0)(0)
    Set qd = db.QueryDefs(strQueryName)
    getSQL = qd.sql

GetSQL_Exit:
    On Error Resume Next
    qd.Close
    Exit Function

GetSQL_Err:
    Resume GetSQL_Exit

End Function


Function GetSysDir() As String
'//============================================================//
'/|   FUNCTION:  GetSysDir                                     |/
'/| PARAMETERS:  -NONE-                                        |/
'/|    RETURNS:  i.e. C:\WIN95\SYSTEM                          |/
'/|    PURPOSE:  String, Path to the Windows System Directory  |/
'/|      USAGE:  s = GetSysDir()                               |/
'/|         BY:  Sean                                          |/
'/|       DATE:  11/30/96                                      |/
'/|    HISTORY:                                                |/
'//============================================================//

Dim lpBuffer As String * 255
Dim Length As Long

   Length = GetSystemDirectory(lpBuffer, Len(lpBuffer))
   GetSysDir = Left(lpBuffer, Length)

End Function


Function GetSysVersions32()
'//================================================================================//
'/|   FUNCTION: GetSysVersions32                                                   |/
'/| PARAMETERS: -NONE-                                                             |/
'/|    RETURNS: -NOTHING-                                                          |/
'/|    PURPOSE: Get the Version and Build of the Operating System                  |/
'/|      USAGE: l = GetSysVersions32()                                             |/
'/|         BY: Sean                                                               |/
'/|       DATE: 11/30/96                                                           |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim v As typOSVERSIONINFO, RetVal As Long
Dim WindowsVersion As String, BuildVersion As String
Dim PlatformName As String

   v.dwOSVersionInfoSize = Len(v)
   RetVal = GetVersionEx(v)

   WindowsVersion = v.dwMajorVersion & "." & v.dwMinorVersion
   BuildVersion = v.dwBuildNumber And &HFFFF&

   Select Case v.dwPlatformId
   Case VER_PLATFORM_WIN32_WINDOWS
      PlatformName = "Windows 95"
   Case VER_PLATFORM_WIN32_NT
      PlatformName = "Windows NT"
   End Select

   MsgBox "Platform: " & PlatformName & vbCrLf & "Version: " & WindowsVersion & vbCrLf & "Build: " & BuildVersion

End Function


Public Function GetScalarValue(ByVal sql As String, Optional fieldName As String = "0") As String
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler
    
    If VBA.InStr(1, sql, "SELECT", vbTextCompare) > 0 Then
        GetScalarValue = Application.CurrentProject.Connection.Execute(sql).Fields(VBA.IIf(fieldName = "0", 0, fieldName)).value
    Else
        
    End If

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox "Error in: basUtility.GetScalarValue" & vbNewLine & Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function


Function GetTextFromWeb(sURL As String) As String
'//========================================================//
'/|   FUNCTION:  GetTextFromWebURL                         |/
'/| PARAMETERS:  sURL, the URL to get the Web Text From    |/
'/|    RETURNS:  string of data                            |/
'/|    PURPOSE:  Get the data download from Web            |/
'/|      USAGE:  s = GetTextFromWeb("http://www.nike.com") |/
'/|         BY:  Sean                                      |/
'/|       DATE:  11/12/2013                                |/
'/|    HISTORY:                                            |/
'//========================================================//
Dim oHTTp As Object
Dim s As String
's   = GetTextFromWeb("http://prwolfe.bol.ucla.edu/cfootball/scores.htm")
'str = GetTextFromWeb("http://colleges.usnews.rankingsandreviews.com/best-colleges/rankings/national-universities/data/spp+300")

'''''''YOU CAN LOOK THIS METHOD UP -- IT'S REALLY SLOW, and FLAKY
'''''' Set ie = CreateObject("InternetExplorer.Application")
'''''' sDOC = objDoc.body.innerText  (or use innerHTML here)

   Set oHTTp = CreateObject("Microsoft.XMLHTTP")
   oHTTp.Open "GET", sURL, False                         'can replace "GET" with "POST" for some other option??
   oHTTp.Send
   s = oHTTp.responseText
   s = StripHTML(s)
   GetTextFromWeb = s
   'Debug.Print s
   Set oHTTp = Nothing
End Function


Public Function GetVersion() As String
'**********************************************
'Author  :  RKP
'Date/Ver:  03-12-13/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

'    Dim version     As String
'    Dim frm         As Form_MainPage
'    Dim ctr         As Integer
    
    If TableExists("tsysSettings") Then
       GetVersion = "v" & basUtility.GetScalarValue("SELECT KeyValue FROM tsysSettings WHERE KeyName = 'Version'")
    Else
       GetVersion = "v" & basUtility.GetSetting("Version", "tsysSettings")
    End If
    'Set frm = Forms("MainPage")
    'frm.lblVersion = version
'    Set frm = New Form_MainPage
'    frm.lblVersion.Caption = "v" & version
'
'    frm.Refresh
'    frm.Repaint
    
'    For Each frm In Application.Forms
'        Debug.Print frm.Name
'        'frm.Controls("lblVersion").value = version
'        For ctr = 0 To frm.Controls.count - 1
'            If frm.Controls.Item(ctr).value = "lblVersion" Then
'                Debug.Print frm.Controls.Item(ctr).value
'            End If
'        Next
'    Next

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function


Function GetWinDir() As String
'//=====================================================//
'/|   FUNCTION:  GetWinDir                              |/
'/| PARAMETERS:  -NONE-                                 |/
'/|    RETURNS:  C:\WIN95                               |/
'/|    PURPOSE:  String, Path to the Windows directory  |/
'/|      USAGE:  s = GetWinDir()                        |/
'/|         BY:  Sean                                   |/
'/|       DATE:  11/30/96                               |/
'/|    HISTORY:                                         |/
'//=====================================================//
  
Dim lpBuffer As String * 255
Dim Length As Long

   Length = GetWindowsDirectory(lpBuffer, Len(lpBuffer))
   GetWinDir = Left(lpBuffer, Length)

End Function


Function HASHER(strIN As String) As String
'//=====================================================//
'/|   FUNCTION: HASHER()                                |/
'/| PARAMETERS: strIN                                   |/
'/|    RETURNS: Hashed HEXIDECIMAL X of LONG INT crc32 Y|/
'/|      USAGE: HashedFAC = HASHER(sFacilityName)       |/
'/|         BY: Sean                                    |/
'/|       DATE: 1/20/2021                               |/
'/|    HISTORY:                                         |/
'//=====================================================//
   
   HASHER = ""
   HASHER = Hex(Crc32(StrConv(strIN, vbFromUnicode)))  ''// AKA  Hex(Crc32(StrConv([COL_NAME],128))) AS HASHER
                                                            ''// THE 128 IN THERE IS 'vbFromUnicode' or vice-versa
   'HASHER("Sean T. MacDermant")  =  5A525D0E
   'e.g.  crc32(strconv("Sean T. MacDermant", 128))  =  1515347214  &  Hex(1515347214)  =  5A525D0E

   'SEE Query1
   'SELECT tlkColorCrayola.ID, tlkColorCrayola.CrayolaColor, HASHER([CrayolaColor]) AS HASHCrayColor FROM tlkColorCrayola;

End Function


Function HideColumns(formname As String, subformname As String, fieldName As String)
   Forms(formname).Form(subformname).Controls(fieldName).ColumnHidden = True
End Function


Public Function HideNavigationPane()
   On Error Resume Next
   DoCmd.NavigateTo "acNavigationCategoryObjectType"
   DoCmd.RunCommand acCmdWindowHide
End Function


Public Function HideRibbon()
   On Error Resume Next
   DoCmd.ShowToolbar "Ribbon", acToolbarNo
End Function


Public Function HowManyNulls(ByVal pTable As String) As Long
   'HowManyNulls("tlkMCUST")   'HowManyNulls("tlkMPROD")
    Dim db As dao.Database
    Dim fld As dao.Field
    Dim tdf As dao.TableDef
    Dim lngNulls As Long
    Dim lngTotal As Long
    HowManyNulls = 0
    Set db = CurrentDb
    Set tdf = db.TableDefs(pTable)
    For Each fld In tdf.Fields
        'lngNulls = DCount("*", pTable, fld.Name & " Is Null")
        ' accommodate field names which need bracketing ...
        lngNulls = DCount("*", pTable, "[" & fld.name & "] Is Null")
        lngTotal = lngTotal + lngNulls
        If lngNulls >= 2 Then
         Debug.Print fld.name, Space(3), lngNulls
        End If
    Next fld
    HowManyNulls = lngTotal
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Function


Public Function InvokeLiteWebService( _
    ByVal vsUrl As String, _
    Optional ByVal vvRequest As Variant, _
    Optional ByVal vsVerb As String = "GET" _
) As Object  'MSXML.DOMDocument
'**********************************************
'Author  :  Ravi Poluri
'Date/Ver:  11-28-03/v7.0.113
'Input   :
'Output  :
'Comments:
'A detailed description of the foundation for this AJAX-style (AVAX, in this case) approach can be read at:
'http://developer.mozilla.org/en/docs/AJAX:Getting_Started
'AJAX: Asynchronous JavaScript And XML
'AVAX: Asynchronous VBScript And XML
'**********************************************
   On Error GoTo Err_Handler

   Dim loXmlHttpRequest As Object 'New MSXML.XMLHTTPRequest
   Dim loXmlDocument As Object 'MSXML.DOMDocument
   
   Set loXmlHttpRequest = CreateObject("MSXML2.XMLHTTP")
   Set loXmlDocument = CreateObject("MSXML2.DomDocument")

   'Connect to the ASP.NET URL and send the request
   loXmlHttpRequest.Open vsVerb, vsUrl, varAsync:=False
   If IsMissing(vvRequest) Then
      loXmlHttpRequest.Send
   Else
      loXmlHttpRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
      loXmlHttpRequest.Send vvRequest
   End If

   'Check for an <html> response - that would indicate an error
   'probably occured in ASP.NET and an HTML page was returned with
   'an error message.
   If (loXmlHttpRequest.status <> 200) Or (0 = StrComp("<html>", Mid$(loXmlHttpRequest.responseText, 1, 6), _
   vbTextCompare)) Then
      'Note, this could be enhanced to save the HTML to a file and
      ' open in a browser, or to parse out the text to display nicely...
      Err.Raise vbObjectError + 1000, vsUrl, "An error occurred." _
          & vbCrLf & loXmlHttpRequest.responseText
   Else
      'Retrieve the response (an XML document object) and return it
      Set loXmlDocument = loXmlHttpRequest.responseXML
      Set InvokeLiteWebService = loXmlDocument
   End If

Err_Handler:
   If Err Then
      If Err.Number = -2146697211 Then
'         MsgBox "The web server is not responding to a request for data." & vbNewLine & vbNewLine & "Please try the operation again or report the issue with ""Internet Services"" if the problem persists.", vbExclamation, Application.ActiveWorkbook.Name
      Else
         'Resume
'         MsgBox "An error has occured while accessing data via Web Service." & vbNewLine & vbNewLine & "Error description:" & vbNewLine & Err.Number & " - " & Err.Description, vbExclamation, Application.ActiveWorkbook.Name
      End If
   End If
End Function


Function IsAlpha(strTest As String)
'//=====================================================//
'/|   FUNCTION:  IsAlpha                                |/
'/| PARAMETERS:  strTest, string to Test                |/
'/|    RETURNS:  True if Alpha, False if not            |/
'/|    PURPOSE:  Checks the Leftmost character to see   |/
'/|              if it is an Alpha Char                 |/
'/|      USAGE:  IsAlpha("B210Z")                       |/
'/|         BY:  Sean                                   |/
'/|       DATE:  11/30/96                               |/
'/|    HISTORY:                                         |/
'//=====================================================//

Dim strChar As String
    strChar = UCase$(Left$(strTest, 1))
    IsAlpha = ((Asc(strChar) > 64) And (Asc(strChar) < 91))

End Function


Function IsCurrentUserInGroup(GName As String) As Integer
'//================================================================================//
'/|   FUNCTION: IsCurrentUserInGroup                                               |/
'/| PARAMETERS: GName, a string identifying the group name                         |/
'/|    RETURNS: True if the Current User is in the Group, Otherwise False          |/
'/|    PURPOSE: Checks if the user account name is in the specified group          |/
'/|             Uses a secure workgroup, so it will run for non-priviledged users  |/
'/|      USAGE: IsCurrentUserInGroup("Admins")                                     |/
'/|         BY: Sean                                                               |/
'/|       DATE: 11/30/96                                                           |/
'/|    HISTORY:                                                                    |/
'//================================================================================//'

Dim W As Workspace, u As User, i As Integer
   
   IsCurrentUserInGroup = False
   Set W = DBEngine.CreateWorkspace("", "Admin", "")       '// 2nd=Admins level user name, 3rd=Password

   Set u = W.Users(CurrentUser())
   
   For i = 0 To u.Groups.count - 1
      If u.Groups(i).name = GName Then
         IsCurrentUserInGroup = True
         Exit Function
      End If
   Next i

End Function


Function IsLoaded2(intObjType As Integer, strObjName As String)
'//===============================================================================//
'/|   FUNCTION:  IsLoaded2                                                        |/
'/| PARAMETERS:  intObjType, the Type of Object to Check (A_FORM or A_REPORT)     |/
'/|              strObjName, the Name of the Object to Check                      |/
'/|    RETURNS:  True if Form or Report is Open, False if it is not               |/
'/|    PURPOSE:  Tells if the specified form or report is open in any             |/
'/|              way (Form View, Report View, Datasheet View or Design View)      |/
'/|      USAGE:  IsLoaded2(A_FORM,"Form1")                                        |/
'/|         BY:  Sean                                                             |/
'/|       DATE:  11/30/96                                                         |/
'/|    HISTORY:                                                                   |/
'//===============================================================================//

On Error Resume Next

    IsLoaded2 = SysCmd(SYSCMD_GETOBJECTSTATE, intObjType, strObjName) > 0

End Function

Function IsTableQuery(DbName As String, TName As String) As Integer
'//===============================================================================//
'/|   FUNCTION:  IsTableQuery                                                     |/
'/| PARAMETERS:  DbName, The name of the database.  If null CurrentDB() is used.  |/
'/|              TName, The name of a table or query.                             |/
'/|    RETURNS:  True (it exists) or False (it does not exist).                   |/
'/|    PURPOSE:  Determine if a table or query exists.                            |/
'/|      USAGE:  IsTableQuery("","ThreeKey")                                      |/
'/|         BY:  Sean                                                             |/
'/|       DATE:  11/30/96                                                         |/
'/|    HISTORY:                                                                   |/
'//===============================================================================//

On Error Resume Next
Dim db As Database, Found As Integer, Test As String
Const NAME_NOT_IN_COLLECTION = 3265

   Found = False                                   '// Assume the table or query does not exist.

   If Trim$(DbName) = "" Then                      '// If the database name is empty...
      Set db = CurrentDb()                         '// ...then set Db to the current Db.
   Else
      Set db = DBEngine.Workspaces(0).OpenDatabase(DbName)     '// Otherwise, set Db to the specified open database.
   
      If Err Then                                              '// See if an error occurred.
         MsgBox "Could not find database to open: " & DbName
         IsTableQuery = False
         Exit Function
      End If
   End If
   
   Test = db.TableDefs(TName).name                             '// See if the name is in the Tables collection.
   If Err <> NAME_NOT_IN_COLLECTION Then Found = True
   
   Err = 0
   
   Test = db.QueryDefs(TName$).name                            '// See if the name is in the Queries collection.
   If Err <> NAME_NOT_IN_COLLECTION Then Found = True
   
   db.Close
   
   IsTableQuery = Found

End Function


Function IsUserInGroup(UName As String, GName As String) As Integer
'//================================================================================//
'/|   FUNCTION: IsUserInGroup                                                      |/
'/| PARAMETERS: UName, a string identifying the user name                          |/
'/|             GName, a string identifying the group name                         |/
'/|    RETURNS: True if the User is in the Group, Otherwise False                  |/
'/|    PURPOSE: Checks whether the user name specified belongs to the group        |/
'/|             name specified.  This is specific to the current SYSTEM.MDA only.  |/
'/|      USAGE: IsUserInGroup("Admin", "Admins")                                   |/
'/|         BY: Sean                                                               |/
'/|       DATE: 11/30/96                                                           |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim W As Workspace, u As User, i As Integer, Found As Integer
   
   IsUserInGroup = False
   Set W = DBEngine.CreateWorkspace("", "Admin", "")   '// 2nd=Admins level user name, 3rd=Password

   Found = False                                       '// Checks if user name is valid for this system.mda
   For i = 0 To W.Users.count - 1
      If W.Users(i).name = UName Then
         Found = True
         Set u = W.Users(i)
         Exit For
      End If
   Next i
   If Not Found Then Exit Function

   For i = 0 To u.Groups.count - 1                     '// Check if user in the group
      If u.Groups(i).name = GName Then
         IsUserInGroup = True
         Exit Function
      End If
   Next i
End Function


Public Sub LaunchAssociatedFile(ByVal vsFile As String)
'**********************************************
'Proc#   :  149
'Author  :  RKP
'Date    :  04/09/01
'Input   :  vsFile - Complete path to the file that you want to launch
'Output  :  sMsg - Error Message that may have occured while trying
'           to launch the file, vsFile.
'           rlRetCode - Return Code by lStartDoc().
'Comments:
'You can use the Windows API ShellExecute() function to start the application
'associated with a given document extension without knowing the name of the
'associated application.
'For example, you could start the Paintbrush program by passing the
'filename ARCADE.BMP to the ShellExecute() function.

'Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                   (ByVal hwnd As Long, ByVal lpszOp As String, _
                    ByVal lpszFile As String, ByVal lpszParams As String, _
                    ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                    As Long
'The following table provides descriptions for each parameter:
'Parameter Description
'----------------------------------------------------------------------------
'hwnd          Identifies the parent window. This window receives any
'              message boxes an application produces (for example, for error
'              reporting).
'
'lpszOp        Points to a null-terminated string specifying the operation
'              to perform. This string can be "open" or "print." If this
'              parameter is NULL, "open" is the default value.
'
'lpszFile      Points to a null-terminated string specifying the file
'              to open.
'
'lpszParams    Points to a null-terminated string specifying parameters
'              passed to the application when the lpszFile parameter
'              specifies an executable file. If lpszFile points to a string
'              specifying a document file, this parameter is NULL.
'
'LpszDir       Points to a null-terminated string specifying the default
'              directory.
'
'FsShowCmd     Specifies whether the application window is to be shown when
'              the application is opened.

'The Windows API ShellExecute() function is different from the
'Visual Basic Shell() function in that you can pass the ShellExecute()
'function the name of a document and it will launch the associated application,
'and then pass the filename to the application.

'If the function succeeds, the return value is the instance handle of the application that was run. If there was an error, the return value is less than or equal to 32.
'The file specified by the lpszFile parameter can be a document file or an executable file. If it is a document file, this function opens or prints it depending on the value of the lpszOp parameter. If it is an executable file, this function opens it even if the lpszOp specifies "PRINT."
'**********************************************
   On Error GoTo Err_Handler

   Dim lRet       As Long
   Dim lHandle    As Long
   Dim sMsg       As String
   
   sMsg = ""
   
   If vsFile = "" Then
      GoTo Err_Handler
   End If
   
   'lRet = StartDoc("C:\WINDOWS\ARCADE.BMP")
   'lRet = lStartDoc(vsFile)
   lHandle = GetDesktopWindow()
   lRet = ShellExecute(lHandle, "Open", vsFile, "", "C:\", SW_SHOWNORMAL)
   
   If lRet <= 32 Then
       'There was an error
       Select Case lRet
           Case SE_ERR_FNF
               sMsg = "File not found"
           Case SE_ERR_PNF
               sMsg = "Path not found"
           Case SE_ERR_ACCESSDENIED
               sMsg = "Access denied"
           Case SE_ERR_OOM
               sMsg = "Out of memory"
           Case SE_ERR_DLLNOTFOUND
               sMsg = "DLL not found"
           Case SE_ERR_SHARE
               sMsg = "A sharing violation occurred"
           Case SE_ERR_ASSOCINCOMPLETE
               sMsg = "Incomplete or invalid file association"
           Case SE_ERR_DDETIMEOUT
               sMsg = "DDE Time out"
           Case SE_ERR_DDEFAIL
               sMsg = "DDE transaction failed"
           Case SE_ERR_DDEBUSY
               sMsg = "DDE busy"
           Case SE_ERR_NOASSOC
               sMsg = "No association for file extension"
           Case ERROR_BAD_FORMAT
               sMsg = "Invalid EXE file or error in EXE image"
           Case Else
               sMsg = "Unknown error"
       End Select
       'MsgBox sMsg
   End If

Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", "LaunchAssociatedFile"
   ElseIf lRet <= 32 Then
      'ProcessMsg lRet, sMsg, "", "LaunchAssociatedFile"
   End If
End Sub


Sub LinkAll(sServer As String, sDatabase As String, Optional sUser As String, Optional sPwd As String)
'ONLY FOR SQL Server for now...
'e.g. Call LinkAll("S02ABIPOC3",  "Area52",         "sa", "Memphis123")
'**e.g. Call LinkAll("S02ASQLP388", "FiberOptCentral")  'this one will use Windows Authentication
   
   Dim sSQL     As String
   Dim sConnect As String
   Dim rs       As dao.Recordset
   Dim i        As Integer
   
   sSQL = "SELECT [name] AS TableName " & vbCrLf
   sSQL = sSQL & "  ,[schema_id] " & vbCrLf
   sSQL = sSQL & "  ,[type] " & vbCrLf
   sSQL = sSQL & "  ,[type_desc]" & vbCrLf
   sSQL = sSQL & "FROM [Sys].[objects]" & vbCrLf
   sSQL = sSQL & "WHERE [type] = 'U' OR [type] = 'V'"
   
   
   'CREATE CONNECT STRING
   If Len(sUser) = 0 Then
      '//Use trusted authentication if stUsername is not supplied.
      sConnect = "ODBC;DRIVER=SQL Server;SERVER=" & sServer & ";DATABASE=" & sDatabase & ";Trusted_Connection=Yes"
    Else
      '//WARNING: This will save the username and the password with the linked table information.
      sConnect = "ODBC;DRIVER=SQL Server;SERVER=" & sServer & ";DATABASE=" & sDatabase & ";UID=" & sUser & ";PWD=" & sPwd
    End If

   If IsTableQuery("", "qsysSQLViewObjects") Then
      CurrentDb.QueryDefs.DELETE "qsysSQLViewObjects"
   End If
   
   
   CreateSPTQuery "qsysSQLViewObjects", sConnect, sSQL


   sSQL = "SELECT qsysSQLViewObjects.* " & _
          "FROM qsysSQLViewObjects " & _
          "ORDER BY [type] "

   
   Set rs = CurrentDb.OpenRecordset(sSQL)
   rs.MoveLast
   rs.MoveFirst
   
   While rs.EOF = False
      DoEvents
      i = AttachDSNLessTable("dbo_" & rs!tableName, "dbo." & rs!tableName, sServer, sDatabase)
      rs.MoveNext
   Wend

   Debug.Print "LinkAll Complete."
        ''AttachDSNLessTable("dbo_tlkFacility", "dbo.tlkFacility", "S02ABIPOC3", "Area52", "sa", "Memphis123")
        ''AttachDSNLessTable("dbo_FOC_tlkVendor","dbo.FOC_tlkVendor","S02ASQLP388","FiberOptCentral")

End Sub


Public Function ListFiles(strPath As String, Optional strFileSpec As String, _
    Optional bIncludeSubfolders As Boolean, Optional lst As ListBox)
On Error GoTo Err_Handler
    'Purpose:   List the files in the path.
    'Arguments: strPath = the path to search.
    '           strFileSpec = "*.*" unless you specify differently.
    '           bIncludeSubfolders: If True, returns results from subdirectories of strPath as well.
    '           lst: if you pass in a list box, items are added to it. If not, files are listed to immediate window.
    '               The list box must have its Row Source Type property set to Value List.
    'Method:    FilDir() adds items to a collection, calling itself recursively for subfolders.
    Dim colDirList As New Collection
    Dim varItem As Variant
    
    Call FillDir(colDirList, strPath, strFileSpec, bIncludeSubfolders)
    
    
    'Add the files to a list box if one was passed in. Otherwise list to the Immediate Window.
    If lst Is Nothing Then
        For Each varItem In colDirList
            ListFiles = varItem
        Next
    Else
        For Each varItem In colDirList
        lst.AddItem varItem
        Next
    End If
Exit_Handler:
    Exit Function
Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume Exit_Handler
End Function


Public Function Log(ByVal text As String, Optional folderPath, Optional FileName, Optional recreate)
'**********************************************
'Author  :  RKP
'Date/Ver:  03-23-09/V01
'Input   :
'Output  :
'Comments:
'Recorded on Thu20Jan11, to make it a generic all-purpose function.
'**********************************************
   On Error GoTo Err_Handler

   Dim oFSO          As Object
   Dim oFile         As Object
   Dim writeFlag     As Integer
   Dim FilePath      As String

   Const ForReading = 1, ForWriting = 2, ForAppending = 8

   Set oFSO = CreateObject("Scripting.FileSystemObject")

   FilePath = ""

'   If Sheets("Main").Range("FilePathRunLog").Value = "" Then
      If VBA.IsMissing(folderPath) Then
         folderPath = Application.CurrentProject.Path  'Application.ThisWorkbook.Path  'basUtility.GetFilePath_WorkFolder
      End If
      If VBA.IsMissing(FileName) Then
         FileName = "Run.log"
      End If
      FilePath = folderPath & "\" & FileName
'   Else
'      filePath = Sheets("Main").Range("FilePathRunLog").Value
'   End If

   If Not VBA.IsMissing(recreate) Then
      If VBA.CBool(recreate) Then
         'basUtility.LogCreate folderPath, fileName
         Call oFSO.CreateTextFile(FilePath, True)
      End If
   End If

   If Not oFSO.FileExists(FilePath) Then
      'basUtility.LogCreate folderPath, fileName
      Call oFSO.CreateTextFile(FilePath, True)
   End If

   If text <> "" Then
      Set oFile = oFSO.OpenTextFile(FilePath, ForAppending, False)
      oFile.WriteLine """" & VBA.Now & """,""" & VBA.Replace(text, """", "'") & """"
      'oFile.writeline text
      'oFile.writeline ""
      oFile.Close
   End If

Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", ""
      'MsgBox Err.Number & " - " & Err.Description
      Resume Next
   End If
   Set oFile = Nothing
   Set oFSO = Nothing

   Exit Function
   Resume
End Function


Function MaximizeAccess()
   Dim Maxit As Integer
   Maxit = ShowWindow(hWndAccessApp, SW_SHOWMAXIMIZED)
End Function


Function MinimizeAccess()
   Dim Minit As Integer
   Minit = ShowWindow(hWndAccessApp, SW_SHOWMINIMIZED)
End Function


Public Sub ModuleList()
   Dim dbs As Database
   Dim doc As Document
   Set dbs = CurrentDb
   
   With dbs.Containers!Modules
      For Each doc In .Documents
         ModuleListAllProcs doc.name
      Next
   End With
   
   dbs.Close
   Set dbs = Nothing
   MsgBox "Done"
End Sub


Function ModuleListAllProcs(strModuleName As String)
   Dim mdl As Module
   Dim lngCount As Long, lngCountDecl As Long, lngI As Long
   Dim strProcName As String, astrProcNames() As String
   Dim intI As Integer, strMsg As String
   Dim lngR As Long
   
   ' Open specified Module object.
   DoCmd.OpenModule strModuleName
   ' Return reference to Module object.
   Set mdl = Modules(strModuleName)
   ' Count lines in module.
   lngCount = mdl.CountOfLines
   ' Count lines in Declaration section in module.
   
   lngCountDecl = mdl.CountOfDeclarationLines
   ' Determine name of first procedure.
   strProcName = mdl.ProcOfLine(lngCountDecl + 1, lngR)
   ' Initialize counter variable.
   intI = 0
   ' Redimension array.
   ReDim Preserve astrProcNames(intI)
   ' Store name of first procedure in array.
   astrProcNames(intI) = strProcName
   ' Determine procedure name for each line after declarations.
   For lngI = lngCountDecl + 1 To lngCount
      ' Compare procedure name with ProcOfLine property value.
      
      If strProcName <> mdl.ProcOfLine(lngI, lngR) Then
         ' Increment counter.
         intI = intI + 1
         strProcName = mdl.ProcOfLine(lngI, lngR)
         ReDim Preserve astrProcNames(intI)
         ' Assign unique procedure names to array.
         astrProcNames(intI) = strProcName
      End If
   Next lngI
      
   strMsg = "Procedures in module '" & strModuleName & "': " & vbCrLf & vbCrLf
      
   For intI = 0 To UBound(astrProcNames)
      strMsg = strMsg & astrProcNames(intI) & vbCrLf
   Next intI
   
   ' Dialog box listing all procedures in module.
   MsgBox strMsg
   Debug.Print "MODULE:  " & strModuleName & CRLF & strMsg
End Function


Function Null2Zero(AValue)
   ' Purpose: Return the value 0 if AValue is Null.
   If IsNull(AValue) Then
      Null2Zero = 0
   Else
      Null2Zero = AValue
   End If
End Function


Function ODBCErrorTrap(Err_Num As Integer, Err_Str As String) As Integer
'//=====================================================//
'/|   FUNCTION:  ODBCErrorTrap                          |/
'/| PARAMETERS:  Err_Num, Err_Str, number and string    |/
'/|    RETURNS:  -NONE-                                 |/
'/|    PURPOSE:  Display ODBC Error in a Message Box    |/
'/|      USAGE:  ODBCErrorTrap(intERR, Error(intERR))   |/
'/|         BY:  Sean                                   |/
'/|       DATE:  11/30/96                               |/
'/|    HISTORY:                                         |/
'//=====================================================//

Dim SQL_Err As String
ODBCErrorTrap = False
Beep
If InStr(1, "3146, 3155, 3156, 3157", Err_Num) Then

    SQL_Err = ExtractSQLErr(Err_Num, Err_Str)
    If Not IsNumeric(Left(SQL_Err, 4)) Then
    Exit Function                    '// Exit without displaying the Error (it was not a number)
    Else
    Select Case Left(SQL_Err, 4)
        Case 2601
        MsgBox "Could not Add or Update this Record because" & CRLF() & CRLF() & "There is already a Record with that Primary Key", MB_ICONEXCLAMATION, "Error #" & Left(SQL_Err, 4) & " reported from SQL-Server"
        Case Else
        MsgBox Mid$(SQL_Err, 7), MB_ICONEXCLAMATION, "Error #" & Left(SQL_Err, 4) & " reported from SQL-Server"
    End Select
    End If
Else
    If Err <> 0 Then MsgBox Err & " " & Error$, MB_ICONEXCLAMATION, "Jet Database Engine Error"
End If
ODBCErrorTrap = True

End Function


Public Function Outlook_SendMail( _
   ByVal emailTo As String, _
   ByVal emailCC As String, _
   ByVal emailBCC As String, _
   ByVal emailSubject As String, _
   ByVal emailBody As String, _
   Optional returnReceipt = False _
)
'**********************************************
'Author  :  RKP
'Date/Ver:  07-29-2010/V01
'Input   :
'Output  :
'Comments:
'**********************************************
   On Error GoTo Err_Handler

   Dim bStarted      As Boolean
   Dim oOutlookApp   As Object  'Outlook.Application
   Dim oItem         As Object  'Outlook.MailItem
   
   'Get a reference to Outlook if it is already running
   Set oOutlookApp = GetObject(, "Outlook.Application")

   If oOutlookApp Is Nothing Then
      Set oOutlookApp = CreateObject("Outlook.Application")
      oOutlookApp.session.Logon
      bStarted = True
   End If
   
   'Create a new mail item
   Set oItem = oOutlookApp.CreateItem(0) 'olMailItem)
   
   With oItem
      'Set the recipient for the new email
      .To = emailTo
      If emailCC <> "" Then
         .CC = emailCC
      End If
      If emailBCC <> "" Then
         .BCC = emailBCC
      End If
      .subject = emailSubject
      .Body = emailBody
      .ReadReceiptRequested = returnReceipt
      .Send
   End With
   
   If bStarted Then
      oOutlookApp.Quit
   End If

Err_Handler:
   mlLastErr = Err.Number
   msLastErr = Err.Description
   'Function1 = mlLastErr
   If Err Then
      If Err.Number = 49 Then 'Bad DLL calling convention
         mlLastErr = 0
         msLastErr = ""
         Resume Next
      ElseIf Err.Number = 429 Then 'ActiveX component can't create object
         Resume Next
      Else
         'ProcessMsg Err.Number, Err.Description, "", ""
         MsgBox Err.Number & " - " & Err.Description
      End If
   End If
   Set oItem = Nothing
   Set oOutlookApp = Nothing
   
   Exit Function
   Resume
End Function


Function PadLeft(s, ByVal c As String, n As Integer) As String
   '
   ' Adds character C to the left of S to make it right justified
   ' and N characters long
   '
     If Len(c) = 0 Then c = " "
     If n < 1 Then
       PadLeft = ""
     Else
       PadLeft = Right$(String$(n, Left$(c, 1)) & s, n)
     End If
End Function


Function PadRight(s, ByVal c As String, n As Integer) As String
   '
   ' Adds character C to the right of S to make it left justified
   ' and N characters long
   '
     If Len(c) = 0 Then c = " "
     If n < 1 Then
       PadRight = ""
     Else
       PadRight = Left$(s & String$(n, Left$(c, 1)), n)
     End If
End Function


Function QCntr(X) As Long
   Cntr = Cntr + 1
   QCntr = Cntr
End Function


Sub QMaster(strKind As String, strFieldSet As String, strQTitle As String)

Dim strSQL As String
Dim strFLD As String
Dim i As Integer
Dim db As Database
Dim qd As QueryDef

   Set db = DBEngine(0)(0)
   Set qd = db.QueryDefs("Query1")         '// Open the QueryDef Object
   
   strFLD = DLookup("[Description]", "tblMaster", "[Kind] = 'FIELDSET' And [Code]=" & "'" & strFieldSet & "'")

   strSQL = "SELECT DISTINCTROW Kind, Code, Description, SortCode, " & strFLD & " FROM tblMaster" & " WHERE (((tblMaster.Kind) = " & "'" & strKind & "')) ORDER BY tblMaster.Kind, tblMaster.Code, tblMaster.SortCode;"

   i = ChangeSQL("Query1", strSQL)

   If i = 0 Then
      qd.Close
      Debug.Print "Had to Close the Q"
      i = ChangeSQL("Query1", strSQL)
   End If
   
   Debug.Print "Change SQL Worked?" & i
   Debug.Print strSQL

'  qd.Description = strQTitle
   DoCmd.OpenQuery "Query1"

End Sub


Public Function QueryExists(tblName As String) As Boolean
'//========================================================//
'/|   FUNCTION:  QueryExists                               |/
'/| PARAMETERS:  tblName, string name of table looking for |/
'/|    RETURNS:  -True if yes, False if No                 |/
'/|      USAGE:  QueryExists("ThisTable")                  |/
'/|         BY:  Sean                                      |/
'/|       DATE:  04/05/16                                  |/
'/|    HISTORY:                                            |/
'//========================================================//
   
'?QueryExists("q400NSR099MainWithNSR")
    QueryExists = False
    If DCount("[Name]", "MSysObjects", "[Name] = '" & tblName & "'" & " AND " & "[Type] = 5") = 1 Then  'use type = 1 here for Table (see TableExists)
        QueryExists = True
    End If
End Function


Private Function Registry_SubKey(Optional ByVal Section As String = "") As String
   ' Build SubKey from known values
   Registry_SubKey = "Software\" & m_Company & "\" & m_AppName
   If Len(Section) Then
      Registry_SubKey = Registry_SubKey & "\" & Section
   End If
End Function


Function ReplaceStr(TextIn, SearchStr, Replacement, CompMode As Integer)
'  USAGE:  ReplaceStr("The Gettysburg Address", "Getty", "Pitt", 1)
'RETURNS:  The Pittsburg Address

Dim WorkText As String, Pointer As Integer
   
   If IsNull(TextIn) Then
      ReplaceStr = Null
   Else
      WorkText = TextIn
      Pointer = InStr(1, WorkText, SearchStr, CompMode)
      
      Do While Pointer > 0
         WorkText = Left(WorkText, Pointer - 1) & Replacement & Mid(WorkText, Pointer + Len(SearchStr))
         Pointer = InStr(Pointer + Len(Replacement), WorkText, SearchStr, CompMode)
      Loop
      
      ReplaceStr = WorkText
   End If
   
End Function


Function ReplaceString(ByVal SourceString As String, ByVal OriginalString As String, ByVal NewString As String) As String

'Recursive function ReplaceString searches Source string and replaces ALL OCCURRENCES of OriginalString with NewString.
'If a value for NewString is ommitted (or IsEmpty), then all occurrences of OriginalString are removed from the SourceString!
    
    Dim Position As Integer
    
    If SourceString = "" Or IsNull(SourceString) Then
        ReplaceString = SourceString
    Else
        Position = InStr(1, SourceString, OriginalString)
        If Position > 0 Then
            ReplaceString = (Mid$(SourceString, 1, Position - 1) & NewString & ReplaceString(Mid(SourceString, Position + Len(OriginalString)), OriginalString, NewString))
        Else
            ReplaceString = SourceString
        End If
    End If
 
End Function


Function ReportClose_TSB(strReport As String) As Boolean
  ' Comments  : Closes the named report without errors
  ' Parameters: strReport - name of the Report to close
  ' Returns   : True if successful, False otherwise
  '
  On Error GoTo PROC_ERR

  If SysCmd(acSysCmdGetObjectState, acReport, strReport) = acObjStateOpen Then
    DoCmd.Close acReport, strReport
    ReportClose_TSB = True
  Else
    ReportClose_TSB = False
  End If

PROC_EXIT:
  Exit Function

PROC_ERR:
  ReportClose_TSB = False
  Resume PROC_EXIT

End Function


Function ReportHasData_TSB(rptIn As Report) As Boolean
  ' Comments  : determines if the specified report has any data
  ' Parameters: rptIn - handle to the report (call this function from the report's OnOpen event)
  ' Returns   : True if the report's RecordSource has one or more records, False otherwise
  '
  Dim dbsCurrent As Database
  Dim rstTemp As Recordset

  On Error GoTo PROC_ERR
                                        
  Set dbsCurrent = CurrentDb()

  ' Open a recordset on the report's RecordSource property
  Set rstTemp = dbsCurrent.OpenRecordset(rptIn.RecordSource)

  ' If more than one record exists, the report has data
  ReportHasData_TSB = (rstTemp.RecordCount > 0)

PROC_EXIT:
  Exit Function

PROC_ERR:
  ReportHasData_TSB = False
  Resume PROC_EXIT

End Function


Function RestoreAccess()
   Dim Restoreit As Integer
   Restoreit = ShowWindow(hWndAccessApp, SW_SHOWNORMAL)
End Function


Public Function RunQueriesByPrefix(sPre As String, _
                                   Optional bShowProgress As Boolean = False, _
                                   Optional bUseDisplayNames As Boolean = False) As Boolean
                                  
  ' Comments: Runs Queries with names matching a supplied query name prefix
  '
  '
  ' Params  : Prefix of query, Show Progess Bar flag, Use cleaned up Display Names flag
  ' Returns : True if success, False if failure
  '
  '
  ' Dependencies:
  '   Form   : frmProgressBar
  '   Module : clsProgressBar
  '
  ' Modifications:
  '    Date         Initals  Description
  '    10/10/2019   MLM      Added error handling
  '    10/13/2019   MLM      Added progress bar and display names feature
  
    
    Dim sSQL As String
    Dim sQuery As String
    Dim rs As dao.Recordset
    Dim z As Date
    Dim pb As clsProgressBar
    
 On Error GoTo Catch
 
    RunQueriesByPrefix = False
    
    DoCmd.SetWarnings False
    z = Now()
    
    sSQL = "SELECT Name FROM MSysObjects " & _
            "   WHERE (Flags=32 Or Flags=48 Or Flags=64 Or Flags=80) " & _
            "   AND MSysObjects.Name Like '{0}*' " & _
            "   ORDER BY Name"
    sSQL = Replace(sSQL, "{0}", sPre)
    Set rs = CurrentDb.OpenRecordset(sSQL)
    
     
    If bShowProgress Then
       'Get accurate total record count
        rs.MoveLast
        rs.MoveFirst
        
        'Setup Progress Bar
        Set pb = New clsProgressBar
        pb.Init Title:="Queries By Prefix", CanCancel:=False
        pb.TotalTasks = rs.RecordCount
    End If
    
    While Not rs.EOF
        sQuery = rs.Fields("Name").value
        If bShowProgress Then
            'Use cleaned up display names or query name with prefix
            If bUseDisplayNames Then
               pb.TaskName = QueryDisplayName(sPre, sQuery)
            Else
               pb.TaskName = sQuery
            End If
            
            'Run the action query
            DoCmd.OpenQuery sQuery
            
            'Mark the task complete in the progress bar
            pb.TaskComplete
        Else
            'Run the action query
            DoCmd.OpenQuery sQuery
        End If
        Debug.Print "   Run query by prefix RAN:  " & sQuery & Space(98 - Len(sQuery)) & " ... " & Format(Now() - z, "hh:nn:ss")
        rs.MoveNext
    Wend
    RunQueriesByPrefix = True 'Success
    
Finally:
    DoCmd.SetWarnings True
    If bShowProgress Then Set pb = Nothing 'clean up
    Exit Function
    
Catch:
    RunQueriesByPrefix = False 'Failure
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Query: " _
            & sQuery, vbCritical, "RunQueriesByPrefix SQL Error"
    Resume Finally
    Resume
End Function

Private Function SplitCaps(ByRef strIN As String) As String
  'Split a string on Capital characters and add space between words
  
    Dim objRegex As Object
    
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Global = True
        .Pattern = "([a-z])([A-Z])"
        SplitCaps = .Replace(strIN, "$1 $2")
    End With
    
End Function


Private Function FirstCapitalChar(sString As String) As Long
  'Comment: Find the first capital letter in a string
  '         and return the position in the string
  '
    Dim i As Long
    Dim strSub As String
    Dim nAsc As Integer
    
    FirstCapitalChar = 0
    
    'loop over string 1 char at a time
    For i = 1 To Len(sString)
        strSub = Mid(sString, i, 1)
        nAsc = AscW(strSub)
        If nAsc >= 65 And nAsc <= 90 Then 'is upper case
           'found it
           FirstCapitalChar = i
           Exit For
        End If
    Next i
    
End Function


Private Function QueryDisplayName(ByVal sPre, ByVal sQueryName As String) As String
  'Get the Query by Prefix query display name
    
    QueryDisplayName = ""
    
    'check param
    If Trim(sQueryName) = "" Then Exit Function
    
    'remove prefix from display name string
    sQueryName = Replace(sQueryName, sPre, "")
    
    'remove remaining left most string coding characters
    sQueryName = Mid(sQueryName, FirstCapitalChar(sQueryName))
    
    'split into words on Capitals
    QueryDisplayName = SplitCaps(sQueryName)
    
End Function


Public Function sBrowseForFolder( _
   ByVal vsTitle As String, _
   ByVal vlHandle As Long _
   ) As String
'**********************************************
'Proc#   :  168
'Author  :  RKP
'Date    :  05/10/01
'Input   :  vsTitle - String - Title for the window.
'Output  :
'Comments:
'Opens a Treeview control that displays the directories in a computer
'**********************************************
   On Error GoTo Err_Handler

   Dim lIDList       As Long
   Dim sBuffer       As String
   Dim sTitle        As String
   Dim udtBrowseInfo As TYPE_BROWSEINFO
   
   sBrowseForFolder = ""
   sTitle = vsTitle
   With udtBrowseInfo
      .hWndOwner = vlHandle 'Me.hwnd
      .lpszTitle = lstrcat(sTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
   End With
   
   lIDList = SHBrowseForFolder(udtBrowseInfo)
   
   If (lIDList) Then
      sBuffer = VBA.Strings.Space$(MAX_PATH)
      SHGetPathFromIDList lIDList, sBuffer
      sBuffer = sTrimAPIString(sBuffer) 'Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      sBrowseForFolder = sBuffer
   End If

Err_Handler:
   If Err Then
'      ProcessMsg Err.Number, Err.Description, "", "sBrowseForFolder"
   End If
End Function


Public Function SendMail_SMTP( _
    ByVal emailFromName As String, _
    ByVal emailFrom As String, _
    ByVal emailTo As String, _
    ByVal emailCC As String, _
    ByVal emailBCC As String, _
    ByVal emailSubject As String, _
    ByVal emailBody As String, _
    Optional htmlBody As Boolean = False, _
    Optional htmlCreateBody As Boolean = False, _
    Optional attachment1 As String, _
    Optional attachment2 As String, _
    Optional attachment3 As String, _
    Optional returnReceipt = False _
)
'**********************************************
'Author  :  RKP
'Date/Ver:  10-27-11/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    Const cdoSendUsingPort = 2
    Const cdoBasic = 1
    Dim objCDOConfig
    Dim objCDOMessage
    Dim strSch
    
    
    
    
    strSch = "http://schemas.microsoft.com/cdo/configuration/"
    Set objCDOConfig = CreateObject("CDO.Configuration")
    With objCDOConfig.Fields
        .Item(strSch & "sendusing") = cdoSendUsingPort
        .Item(strSch & "smtpserver") = "smtp.ipaper.com"
        ' Only used if SMTP server requires Authentication
        '.Item(strSch & "smtpauthenticate") = cdoBasic
        '.Item(strSch & "sendusername") = "username"
        '.Item(strSch & "sendpassword") = "******"
        .Update
    End With
        
    Set objCDOMessage = CreateObject("CDO.Message")
    With objCDOMessage
        Set .Configuration = objCDOConfig
        .FROM = emailFromName
        .sender = emailFrom
        .To = emailTo
        .subject = emailSubject
        If htmlBody Then
            If htmlCreateBody Then
                .CreateMHTMLBody = attachment1
            Else
                .htmlBody = emailBody
            End If
        Else
            .TextBody = emailBody
        End If
        .Send
    End With
    
    If Err <> 0 Then
        'MsgBox Err.Number
    End If

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            'MsgBox Err.Number & " - " & Err.Description
            basUtility.Log "Error sending SMTP email notification." & vbNewLine & "SendMail_SMTP" & vbNewLine & Err.Number & " - " & Err.Description
        End If
    End If
    Set objCDOMessage = Nothing
    Set objCDOConfig = Nothing
    
    Exit Function
    Resume
End Function


Function SetApplicationTitle(ByVal MyTitle As String)
'//=====================================================//
'/|   FUNCTION:  SetApplicationTitle                    |/
'/| PARAMETERS:  MyTile, string containing App Title    |/
'/|    RETURNS:  -NONE-                                 |/
'/|    PURPOSE:  Set the Application's Tile Bar Text    |/
'/|      USAGE:  SetApplicationTitle("This Title")      |/
'/|         BY:  Sean                                   |/
'/|       DATE:  11/30/96                               |/
'/|    HISTORY:                                         |/
'//=====================================================//
   
   If SetStartupProperty("AppTitle", dbText, MyTitle) Then
      Application.RefreshTitleBar
   Else
      MsgBox "ERROR: Could not set Application Title"
   End If

End Function


Function SetCntrToZero()
   Cntr = 0
End Function


Private Sub SetDAOFieldProperty _
               (fld As dao.Field, _
 stName As String, vValue As Variant, _
               lType As Long)
    Dim PRP As dao.Property
    
    For Each PRP In fld.Properties
        If StrComp(PRP.name, stName, _
         vbBinaryCompare) = 0 Then
            PRP.value = vValue
            Exit For
        End If
        Set PRP = Nothing
    Next PRP
    
    If PRP Is Nothing Then
        Set PRP = fld.CreateProperty(stName, _
         lType, vValue)
        fld.Properties.Append PRP
    End If
End Sub


Public Sub SetColumnHidden(tableName As String, fieldName As String)

    Dim dbs As dao.Database
    Dim fld As dao.Field
    Dim PRP As dao.Property
    Const conErrPropertyNotFound = 3270

    ' Turn off error trapping.
    On Error Resume Next

    Set dbs = CurrentDb
    
    
    ' Set field property.
    Set fld = dbs.TableDefs(tableName).Fields(fieldName)
    fld.Properties("ColumnHidden") = True
    
    ' Error may have occurred when value was set.
    If Err.Number <> 0 Then
        If Err.Number <> conErrPropertyNotFound Then
            On Error GoTo 0
            MsgBox "Couldn't set property 'ColumnHidden' " & _
                   "on field '" & fld.name & "'", vbCritical
        Else
            On Error GoTo 0
            Set PRP = fld.CreateProperty("ColumnHidden", dbLong, True)
            fld.Properties.Append PRP
        End If
    End If
    
    Set PRP = Nothing
    Set fld = Nothing
    Set dbs = Nothing
    
End Sub


Function SetStartupProperty(prpName As String, prpType As Variant, prpValue As Variant) As Integer
'//==============================================================//
'/|   FUNCTION:  SetStartupProperty                              |/
'/| PARAMETERS:  prpName, string, property to set                |/
'/|              prpType, variant,                               |/
'/|              prpValue, variant,                              |/
'/|    RETURNS:  True on Success and False by default or Failure |/
'/|    PURPOSE:  Change a Startup Property                       |/
'/|      USAGE:  SetStartupProperty("AppTitle", dbText, MyTitle) |/
'/|         BY:  Sean                                            |/
'/|       DATE:  11/30/96                                        |/
'/|    HISTORY:                                                  |/
'//==============================================================//

'Called from SetApplicationTitle

Dim db As Database, PRP As Property, ws As Workspace
Const ERROR_PROPNOTFOUND = 3270

   Set db = CurrentDb()

   ' Set the startup property value
   On Error GoTo Err_SetStartupProperty
   db.Properties(prpName) = prpValue
   SetStartupProperty = True

Bye_SetStartupProperty:
   Exit Function

Err_SetStartupProperty:
   Select Case Err
   'Did the property not exist?
   Case ERROR_PROPNOTFOUND
      ' .. then create it.
      Set PRP = db.CreateProperty(prpName, prpType, prpValue)
      db.Properties.Append PRP
      ' .. and try setting it again.
      Resume
   Case Else
      SetStartupProperty = False
      Resume Bye_SetStartupProperty
   End Select
End Function


Public Function ShowNavigationPane()
   On Error Resume Next
   DoCmd.SelectObject acTable, , True
End Function


Public Function ShowRibbon()
   On Error Resume Next
   DoCmd.ShowToolbar "Ribbon", acToolbarYes
End Function


Function SizeAccess()
   
   Dim cX As Long, cY As Long, cHeight As Long
   Dim cWidth As Long, h As Long
   'Get handle to Microsoft Access.
   h = Application.hWndAccessApp

   cX = 70
   cY = 0
   cWidth = 1140
   cHeight = 780

   'Position Microsoft Access.
   SetWindowPos h, HWND_TOP, cX, cY, cWidth, cHeight, SWP_NOZORDER

End Function


Public Function Speak(ByVal text As String)
'**********************************************
'Author  :  RKP
'Date/Ver:  04-24-09/V01
'Input   :
'Output  :
'Comments:
'**********************************************
   On Error GoTo Err_Handler

   If moVoice Is Nothing Then
      Set moVoice = CreateObject("SAPI.SpVoice")
   End If
   moVoice.Speak text

Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", ""
      'MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Function SplitOut(ByVal InputText As String, iPos As Integer, Optional ByVal Delimiter As String) As String
    Dim varRetVal As Variant
    Dim splitstring As Variant
    On Error GoTo SplitOut_Err
    If Len(Delimiter) < 1 Then Delimiter = ""
    'Exit Function
    splitstring = Split(InputText, Delimiter)
    SplitOut = splitstring(iPos - 1)

SplitOut_Done:
  Exit Function

SplitOut_Err:
   Select Case Err
    Case 9 'or 13 'subscript out of range
      SplitOut = ""
    Case Else
      MsgBox "Error: " & Error$ & " " & Err, , "SPLITOUT"
      SplitOut = ""
   End Select
   Resume SplitOut_Done

End Function

Public Function sTrimAPIString(ByVal vsAPIString As String) As String
   On Error GoTo Err_Handler
'************************************************************
'Procedure#    :  93
'Author  :  RKP
'Date    :  May 25, 2000
'Input   :  vsAPIString - String - String that needs to be trimmed.
'Output  :  String - Trimmed String.
'Comments:
'Almost all Win32 API calls require string of a fixed length to be passed in, eg;
'Dim sUserName As String * 128
'When this string is populated with a value, the length is still 128 and all spaces
'are populated with Chr(0) (vbNullChar). This routine removes all Chr(0)'s from vsAPIString and
'returns just the text in vsAPIString.
'sRemoveChr(0)sFromString
'************************************************************
   
   sTrimAPIString = NULL_STRING
   If InStr(1, vsAPIString, vbNullChar, vbTextCompare) > 0 Then
      sTrimAPIString = Trim$(Left$(vsAPIString, InStr(vsAPIString, vbNullChar) - 1))
   Else
      sTrimAPIString = (Trim$(vsAPIString))
   End If
   
Err_Handler:
   If Err Then
      'Log "Error", "GenUtils.sTrimAPIString", Err.Number, Err.Source & ": " & Err.Description
      'ProcessMsg Err.Number, Err.Description, "", "sTrimAPIString"
   End If
End Function


Function StrLen(AVariant) As Integer
   ' Returns the length of a variant or string, returning zero a Null or zero-length string.
   If IsNull(AVariant) Then
      StrLen = 0
   Else
      StrLen = Len(AVariant)
   End If
End Function


Function SwitchMDB(strMDB As String)

On Error GoTo SwitchMDB_Err

'//   gView_Mode = gSHUTDOWN_MODE
'//
'//   SendKeys "{F11}"
'//   SendKeys "%fo" & gsCurrent_Directory & "\" & strMDB & "~", False

SwitchMDB_Done:
   Exit Function

SwitchMDB_Err:
   Exit Function

End Function


Function StripHTML(sHTML As String) As String
'//===============================================================//
'/|   FUNCTION:  StripHTML                                        |/
'/| PARAMETERS:  sHTML, the HTML in a string to convert to text   |/
'/|    RETURNS:  string of data                                   |/
'/|    PURPOSE:  convert HTML to Text or text-like substance      |/
'/|      USAGE:  s = GetTextFromWeb("http://www.nike.com")        |/
'/|         BY:  Sean                                             |/
'/|       DATE:  11/12/2013                                       |/
'/|    HISTORY:                                                   |/
'//===============================================================//
Dim RegEx As Object
Dim sInput As String
Dim sOut As String
Set RegEx = CreateObject("vbscript.regexp")
 
   sInput = sHTML

   With RegEx
      .Global = True
      .IgnoreCase = True
      .MultiLine = True
      .Pattern = "<[^>]+>" 'Regular Expression for HTML Tags.
   End With

   sOut = RegEx.Replace(sInput, "")
   StripHTML = sOut
   Set RegEx = Nothing
End Function


Public Function TableExists(tblName As String) As Boolean
'//========================================================//
'/|   FUNCTION:  TableExists                               |/
'/| PARAMETERS:  tblName, string name of table looking for |/
'/|    RETURNS:  -True if yes, False if No                 |/
'/|      USAGE:  TableExists("ThisTable")                  |/
'/|         BY:  Sean                                      |/
'/|       DATE:  04/05/16                                  |/
'/|    HISTORY:                                            |/
'//========================================================//
   
'?TableExists("q400NSR099MainWithNSR")
    TableExists = False
    If DCount("[Name]", "MSysObjects", "[Name] = '" & tblName & "'" & " AND " & "[Type] = 1") = 1 Then  'use type = 5 here for Query
        TableExists = True
    End If
End Function


Public Function TextFile_ReadAndWrite()
'**********************************************
'Author  :  RKP
'Date/Ver:  11-17-2010/V01
'Input   :
'Output  :
'Comments:
'**********************************************
   On Error GoTo Err_Handler

   Dim fileHandleIn As Integer
   Dim filePathIn As String
   Dim fileHandleOut As Integer
   Dim filePathOut As String
   Dim ctr As Long
   Dim txt As String
   
   filePathIn = "C:\Documents and Settings\rpoluri\Desktop\DTS_HEPPR.txt"
   fileHandleIn = VBA.FreeFile()
   Open filePathIn For Input As #fileHandleIn
   
   filePathOut = "C:\Documents and Settings\rpoluri\Desktop\DTS_HEPPR_DataPullAndLoadConsole.txt"
   fileHandleOut = VBA.FreeFile()
   Open filePathOut For Output As #fileHandleOut
   'Open filePathOut For Append As #fileHandleOut
   
   While Not EOF(fileHandleIn)
      DoEvents
      Line Input #fileHandleIn, txt
      Debug.Print txt
      
      Print #fileHandleOut, "sb.AppendLine(""" & VBA.Replace(txt, """", """""") & " "")"
   Wend
   Close fileHandleIn
   Close fileHandleOut
   
   MsgBox "Finished!"

Err_Handler:
   mlLastErr = Err.Number
   msLastErr = Err.Description
   'Function1 = mlLastErr
   If Err Then
      If Err.Number = 49 Then 'Bad DLL calling convention
         mlLastErr = 0
         msLastErr = ""
         Resume Next
      Else
         'ProcessMsg Err.Number, Err.Description, "", ""
         MsgBox Err.Number & " - " & Err.Description
      End If
   End If
   
   Exit Function
   Resume
End Function


Function TXTImportFromText(strTableName As String, strFileName As String, Optional ByVal strDelim As String = vbTab) As Boolean
   'This function is used to import delimited files, not fixed-width.
   'SAP extracts are pipe-delimited and are identified automatically.

   '//http://saplsmw.com/Import_a_text_file_as_a_table_in_MS_Access_using_VBA
   '//SUBMITTED BY JIMBO ON TUE, 07/12/2011 - 01:01

   '//  ?TXTImportFromText("tableE", "C:\Users\smacder.NAIPAPER\Desktop\Text File Proc\E--KNA1_JackedUpHdr12345PIPEDeLim&HUGE_SAMPLEONLY_5MB.txt", "|")
   '//  ?TXTImportFromText("tableEClean", "C:\Users\smacder.NAIPAPER\Desktop\Text File Proc\E--KNA1_JackedUpHdr12345PIPEDeLim&HUGE_SAMPLEONLY_5MB_CLN.txt", "|")
   
   
    Dim rs As dao.Recordset
    Dim nCurrent As Long, nFieldCount As Long, nRecordCount As Long
    Dim RetVal As Variant, nCurRec As Long, dnow As Date, nCurSec As Long
    Dim nTotalSeconds As Long, nSecondsLeft As Long
    Dim nTotalbytes As Long, nFileLen As Long
    Dim strTest As Variant
    Dim strTemp As String
    Dim strHeadersIn() As String
    Dim strHeaders(999) As String
    Const nReadAhead As Long = 930000
    Dim nSizes(999) As Long, strRecords(nReadAhead) As String, nRecords As Long, nLoaded As Long
    Dim strFields() As String
    
    Dim nHeaders As Long
    Dim isSAP As Boolean
    
    TXTImportFromText = False
    
    nFileLen = FileLen(strFileName)
    RetVal = SysCmd(acSysCmdSetStatus, "Preparing to import " & strTableName & " from " & strFileName & "...")
    RetVal = DoEvents()
    
    Open strFileName For Input As #1
    Line Input #1, strTest
    If Left(strTest, 6) = "Table:" Then 'This is an SAP extract!
        isSAP = True
        Line Input #1, strTest
        Line Input #1, strTest
        Line Input #1, strTest  'Fourth line has the headers!
    Else
        isSAP = False
    End If
    
    If InStr(1, strTest, "|", vbTextCompare) Then
        strDelim = "|"
    End If
    
    nTotalbytes = nTotalbytes + Len(strTest) + 2 ' +2 for vbCrLf--This line prevents div by zero later...
    strTest = Trim(strTest)
    If Right(strTest, 1) = strDelim Then
        strTest = Left(strTest, Len(strTest) - 1)
    End If
    strHeadersIn = Split(Trim(strTest), strDelim)
    nHeaders = 0
    
    
    For Each strTest In strHeadersIn
        nHeaders = nHeaders + 1
        strTest = Replace(Replace(strTest, " ", ""), ".", "")
        strTest = Replace(Replace(strTest, " ", ""), ".", "")
        If Len(Trim(strTest)) = 0 Then
            strHeaders(nHeaders) = "HEADER" & Right("000" & nHeaders, 3)
        Else
            strHeaders(nHeaders) = Trim(strTest)
        End If
        For nCurrent = 1 To nHeaders - 1
            If strHeaders(nHeaders) = strHeaders(nCurrent) Then
                strHeaders(nHeaders) = strHeaders(nHeaders) & nHeaders
            End If
        Next
    Next
    strHeaders(0) = nHeaders
    RetVal = SysCmd(acSysCmdClearStatus)
    RetVal = SysCmd(acSysCmdInitMeter, "Preparing to import " & strTableName & " from " & strFileName & "...", nReadAhead)
    RetVal = DoEvents()
    
    Do While Not EOF(1) And nRecords < nReadAhead 'Read through the file and get the maximum sizes for fields in advance.
        Line Input #1, strTest
        strTest = Trim(strTest)
        If Right(strTest, 1) = strDelim Then
            strTest = Left(strTest, Len(strTest) - 1)
        End If
        If isSAP And Left(strTest, 20) = "--------------------" Then
            strTest = ""  'Skip this line!
        End If
        If Len(strTest) > 0 Then
            nRecords = nRecords + 1
            strRecords(nRecords) = strTest
            strFields = Split(strTest, strDelim)
            nCurrent = 0
            For Each strTest In strFields
                nCurrent = nCurrent + 1
                If Len(strTest) > nSizes(nCurrent) Then
                    nSizes(nCurrent) = Len(strTest)
                End If
            Next
            If Second(Now) <> nCurSec Then
                nCurSec = Second(Now)
                RetVal = SysCmd(acSysCmdUpdateMeter, nRecords)
                RetVal = DoEvents()
            End If
        End If
    Loop
    
    
    If TXTCreateTable(strTableName, strHeaders, nSizes) Then
        If isSAP Then
            For nCurrent = 1 To nHeaders
                If Left(strHeaders(nCurrent), 8) = "HEADER00" Then
                    strHeaders(nCurrent) = ""  'Don't bother importing this field.
                End If
            Next
        End If
        Set rs = CurrentDb.OpenRecordset(strTableName)
        nLoaded = 0
        nTotalSeconds = 0
        Do While Not EOF(1) Or nLoaded < nRecords
            nCurRec = nCurRec + 1
            If Second(Now()) <> nCurSec Then
                nCurSec = Second(Now())
                nTotalSeconds = nTotalSeconds + 1
                'RetVal = DoEvents()
                If nTotalSeconds > 3 Then
                    'nSecondsLeft = Int(((nTotalSeconds / nCurRec) * rs.RecordCount) * ((rs.RecordCount - nCurRec) / rs.RecordCount))
                    nSecondsLeft = Int(((nTotalSeconds / nTotalbytes) * nFileLen) * ((nFileLen - nTotalbytes) / nFileLen))
                    RetVal = SysCmd(acSysCmdRemoveMeter)
                    RetVal = SysCmd(acSysCmdInitMeter, "Importing " & strTableName & " from " & strFileName & "... " & nSecondsLeft & " seconds remaining.", nFileLen)
                    RetVal = SysCmd(acSysCmdUpdateMeter, nTotalbytes)
                    RetVal = DoEvents()
                End If
            End If
            If nLoaded < nRecords Then
                nLoaded = nLoaded + 1
                strTest = strRecords(nLoaded)
            Else
                Line Input #1, strTest
            End If
            nTotalbytes = nTotalbytes + Len(strTest) + 2 'vbCrLf
            strTest = Trim(strTest)
            If Right(strTest, 1) = strDelim Then
                strTest = Left(strTest, Len(strTest) - 1)
            End If
            If isSAP And Left(strTest, 20) = "--------------------" Then
                strTest = ""  'Skip this line!
            End If
            If Len(strTest) > 0 Then
                strFields = Split(strTest, strDelim)
                nCurrent = 0
                rs.AddNew
                For Each strTest In strFields
                    nCurrent = nCurrent + 1
                    If Len(Trim(strHeaders(nCurrent))) > 0 Then
                        rs.Fields(strHeaders(nCurrent)).value = Trim(strFields(nCurrent - 1))
                    End If
                Next
                rs.Update
            End If
        Loop
        rs.Close
    End If
    Close #1
    RetVal = SysCmd(acSysCmdRemoveMeter)
    TXTImportFromText = True
End Function


Function TXTImportFromText2(strTableName As String, strFileName As String, Optional ByVal strDelim As String = vbTab) As Boolean
   'This function is used to import delimited files, not fixed-width.
   'SAP extracts are pipe-delimited and are identified automatically.

   '//http://saplsmw.com/Import_a_text_file_as_a_table_in_MS_Access_using_VBA
   '//SUBMITTED BY JIMBO ON TUE, 07/12/2011 - 01:01

   '//  ?TXTImportFromText2("tableE", "C:\Users\smacder.NAIPAPER\Desktop\Text File Proc\E--KNA1_JackedUpHdr12345PIPEDeLim&HUGE_SAMPLEONLY_5MB.txt", "|")
   '//  ?TXTImportFromText2("tableEClean", "C:\Users\smacder.NAIPAPER\Desktop\Text File Proc\E--KNA1_JackedUpHdr12345PIPEDeLim&HUGE_SAMPLEONLY_5MB_CLN.txt", "|")
   
   '** this doesn't remove spaces or periods
   
    Dim rs As dao.Recordset
    Dim nCurrent As Long, nFieldCount As Long, nRecordCount As Long
    Dim RetVal As Variant, nCurRec As Long, dnow As Date, nCurSec As Long
    Dim nTotalSeconds As Long, nSecondsLeft As Long
    Dim nTotalbytes As Long, nFileLen As Long
    Dim strTest As Variant
    Dim strTemp As String
    Dim strHeadersIn() As String
    Dim strHeaders(999) As String
    Const nReadAhead As Long = 30000
    Dim nSizes(999) As Long, strRecords(nReadAhead) As String, nRecords As Long, nLoaded As Long
    Dim strFields() As String
    
    Dim nHeaders As Long
    Dim isSAP As Boolean
    
    TXTImportFromText2 = False
    
    nFileLen = FileLen(strFileName)
    RetVal = SysCmd(acSysCmdSetStatus, "Preparing to import " & strTableName & " from " & strFileName & "...")
    RetVal = DoEvents()
    
    Open strFileName For Input As #1
    Line Input #1, strTest
    If Left(strTest, 6) = "Table:" Then 'This is an SAP extract!
        isSAP = True
        Line Input #1, strTest
        Line Input #1, strTest
        Line Input #1, strTest  'Fourth line has the headers!
    Else
        isSAP = False
    End If
    
    If InStr(1, strTest, "|", vbTextCompare) Then
        strDelim = "|"
    End If
    
    nTotalbytes = nTotalbytes + Len(strTest) + 2 ' +2 for vbCrLf--This line prevents div by zero later...
    strTest = Trim(strTest)
    If Right(strTest, 1) = strDelim Then
        strTest = Left(strTest, Len(strTest) - 1)
    End If
    strHeadersIn = Split(Trim(strTest), strDelim)
    nHeaders = 0
    
    
    For Each strTest In strHeadersIn
        nHeaders = nHeaders + 1
        ''''strTest = Replace(Replace(strTest, " ", ""), ".", "")
        ''''strTest = Replace(Replace(strTest, " ", ""), ".", "")
        If Len(Trim(strTest)) = 0 Then
            strHeaders(nHeaders) = "HEADER" & Right("000" & nHeaders, 3)
        Else
            strHeaders(nHeaders) = Trim(strTest)
        End If
        For nCurrent = 1 To nHeaders - 1
            If strHeaders(nHeaders) = strHeaders(nCurrent) Then
                strHeaders(nHeaders) = strHeaders(nHeaders) & nHeaders
            End If
        Next
    Next
    strHeaders(0) = nHeaders
    RetVal = SysCmd(acSysCmdClearStatus)
    RetVal = SysCmd(acSysCmdInitMeter, "Preparing to import " & strTableName & " from " & strFileName & "...", nReadAhead)
    RetVal = DoEvents()
    
    Do While Not EOF(1) And nRecords < nReadAhead 'Read through the file and get the maximum sizes for fields in advance.
        Line Input #1, strTest
        strTest = Trim(strTest)
        If Right(strTest, 1) = strDelim Then
            strTest = Left(strTest, Len(strTest) - 1)
        End If
        If isSAP And Left(strTest, 20) = "--------------------" Then
            strTest = ""  'Skip this line!
        End If
        If Len(strTest) > 0 Then
            nRecords = nRecords + 1
            strRecords(nRecords) = strTest
            strFields = Split(strTest, strDelim)
            nCurrent = 0
            For Each strTest In strFields
                nCurrent = nCurrent + 1
                If Len(strTest) > nSizes(nCurrent) Then
                    nSizes(nCurrent) = Len(strTest)
                End If
            Next
            If Second(Now) <> nCurSec Then
                nCurSec = Second(Now)
                RetVal = SysCmd(acSysCmdUpdateMeter, nRecords)
                RetVal = DoEvents()
            End If
        End If
    Loop
    
    
    If TXTCreateTable(strTableName, strHeaders, nSizes) Then
        If isSAP Then
            For nCurrent = 1 To nHeaders
                If Left(strHeaders(nCurrent), 8) = "HEADER00" Then
                    strHeaders(nCurrent) = ""  'Don't bother importing this field.
                End If
            Next
        End If
        Set rs = CurrentDb.OpenRecordset(strTableName)
        nLoaded = 0
        nTotalSeconds = 0
        Do While Not EOF(1) Or nLoaded < nRecords
            nCurRec = nCurRec + 1
            If Second(Now()) <> nCurSec Then
                nCurSec = Second(Now())
                nTotalSeconds = nTotalSeconds + 1
                'RetVal = DoEvents()
                If nTotalSeconds > 3 Then
                    'nSecondsLeft = Int(((nTotalSeconds / nCurRec) * rs.RecordCount) * ((rs.RecordCount - nCurRec) / rs.RecordCount))
                    nSecondsLeft = Int(((nTotalSeconds / nTotalbytes) * nFileLen) * ((nFileLen - nTotalbytes) / nFileLen))
                    RetVal = SysCmd(acSysCmdRemoveMeter)
                    RetVal = SysCmd(acSysCmdInitMeter, "Importing " & strTableName & " from " & strFileName & "... " & nSecondsLeft & " seconds remaining.", nFileLen)
                    RetVal = SysCmd(acSysCmdUpdateMeter, nTotalbytes)
                    RetVal = DoEvents()
                End If
            End If
            If nLoaded < nRecords Then
                nLoaded = nLoaded + 1
                strTest = strRecords(nLoaded)
            Else
                Line Input #1, strTest
            End If
            nTotalbytes = nTotalbytes + Len(strTest) + 2 'vbCrLf
            strTest = Trim(strTest)
            If Right(strTest, 1) = strDelim Then
                strTest = Left(strTest, Len(strTest) - 1)
            End If
            If isSAP And Left(strTest, 20) = "--------------------" Then
                strTest = ""  'Skip this line!
            End If
            If Len(strTest) > 0 Then
                strFields = Split(strTest, strDelim)
                nCurrent = 0
                rs.AddNew
                For Each strTest In strFields
                    nCurrent = nCurrent + 1
                    If Len(Trim(strHeaders(nCurrent))) > 0 Then
                        rs.Fields(strHeaders(nCurrent)).value = Trim(strFields(nCurrent - 1))
                    End If
                Next
                rs.Update
            End If
        Loop
        rs.Close
    End If
    Close #1
    RetVal = SysCmd(acSysCmdRemoveMeter)
    TXTImportFromText2 = True
End Function


Function TXTCreateTable(strTableName As String, strFields() As String, nSizes() As Long) As Boolean
    Dim nCounter As Long
    Dim dbs As dao.Database
    'Now create the database.  Rename the old database if necessary.
    Set dbs = CurrentDb
    Dim tdf As dao.TableDef
    Dim fld1 As dao.Field
    Dim fld2 As dao.Field
    Dim fName As String
    Dim fType As Integer
    Dim fSize As Integer
    
    On Error GoTo ErrorHandler
    'Check for existence of TargetTable
    nCounter = 0
    Do While nCounter < dbs.TableDefs.count
        If dbs.TableDefs(nCounter).name = strTableName Then
            'Delete TargetTable--must start from scratch
            dbs.TableDefs.DELETE (strTableName)
        End If
        nCounter = nCounter + 1
    Loop
    
    Set tdf = dbs.CreateTableDef(strTableName)
    For nCounter = 1 To Val(strFields(0))
        fName = strFields(nCounter)
        fType = dbText  'Never encounter type-mismatch if we always use text.
        fSize = nSizes(nCounter) 'fSize = 255
        Set fld1 = tdf.CreateField(fName, fType, fSize)
        fld1.AllowZeroLength = True
        fld1.Required = False
        tdf.Fields.Append fld1
    Next
    'Create the table in the database
    dbs.TableDefs.Append tdf
    dbs.TableDefs.Refresh
    TXTCreateTable = True
    Exit Function
ErrorHandler:
    MsgBox "Error number " & Err.Number & ": " & Err.Description
    TXTCreateTable = False
    Exit Function
End Function


Function TrailingSlash(varIn As Variant) As String
    If Len(varIn) > 0 Then
        If Right(varIn, 1) = "\" Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & "\"
        End If
    End If
End Function


Function uMAX(a, b)
   'This function calculates the maximum of a and b
   uMAX = -a * (a >= b) - b * (a < b)
End Function


Function uMIN(a, b)
   'calculates minimum of a and b
   uMIN = -a * (a <= b) - b * (a > b)
End Function


Public Function UpdateStatus(ByVal statusText As String)
'**********************************************
'Author  :  RKP
'Date/Ver:  11-30-12/V01
'Input   :
'Output  :
'Comments:
'**********************************************

   'UpdateStatus("Sending Planner Alerts...")
    On Error GoTo Err_Handler

    Application.SysCmd acSysCmdSetStatus, statusText

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Function
    Resume
End Function


Public Function UrlEncode(ByVal vsVal) As String
'**********************************************
'Author  :  Ravi Poluri
'Date/Ver:  11-28-03/v7.0.113
'Input   :
'Output  :
'Comments:
' Encode parameter name or value that will go into the query string of the URL.
'**********************************************
   On Error GoTo Err_Handler
   
   Const lsOK_CHARS$ = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_@.*"
   Dim llIdx As Long
   Dim lsChar As String
   llIdx = 1
   Do
      lsChar = Mid$(vsVal, llIdx, 1)
      'is the current character is a space, change it to a "+"
      If lsChar = " " Then
          Mid$(vsVal, llIdx, 1) = "+"
      'is the current character OK, or does it need to be encoded?
      ElseIf InStr(1, lsOK_CHARS, lsChar) <= 0 Then
          vsVal = Mid$(vsVal, 1, llIdx - 1) & _
              "%" & Right$("0" + Hex(Asc(lsChar)), 2) & Mid$(vsVal, llIdx + 1)
          llIdx = llIdx + 2
      End If
      llIdx = llIdx + 1
   Loop While (llIdx <= Len(vsVal))
   
   UrlEncode = vsVal
   
Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", "UrlEncode"
'      MsgBox Err.Number & " - " & Err.Description
   End If
End Function


Public Sub UrlLog(ByVal vsTaskID As String, Optional millID, Optional pArray)
'**********************************************
'Author  :  RKP
'Date/Ver:  02-06-08/V14
'Input   :
'Output  :
'Comments:
'**********************************************
   On Error GoTo Err_Handler

   Dim SUCCESS    As Boolean

   If VBA.IsMissing(millID) Then
      millID = ""
   End If
   If VBA.IsMissing(pArray) Then
      pArray = ""
   End If
   basUtility.CallWebService "bmosdev", SUCCESS, vsTaskID, millID, pArray
   'basUtility.CallWebService "bmosdev", success, vsTaskID, "", "UNM:" & VBA.Environ("USERNAME") & ";VER:" & Application.ThisWorkbook.VBProject.Description & ";FNM:" & Application.ThisWorkbook.FullName

Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", ""
      'MsgBox Err.Number & " - " & Err.Description
   End If
End Sub


Function uROUND(a, b)
   'rounds the number a to b decimal places
   uROUND = (CLng(a * 10 ^ b)) / (10 ^ b)
End Function


Function SizeName(strName As String, intNameLength As Integer)

If Len(strName) >= intNameLength Then
   SizeName = Left(strName, intNameLength)
Else
   SizeName = strName & Space(intNameLength - Len(strName))
End If

End Function


Sub WaitSecs(secs)
'//=====================================================//
'/|        SUB:  WaitSecs                               |/
'/| PARAMETERS:  secs, the number of seconds to wait    |/
'/|    RETURNS:  -NONE-                                 |/
'/|    PURPOSE:  Wait a number of seconds               |/
'/|      USAGE:  Call WaitSecs(3)                       |/
'/|         BY:  Sean                                   |/
'/|       DATE:  11/30/96                               |/
'/|    HISTORY:                                         |/
'//=====================================================//

Dim sglStart As Single
Dim intTemp As Integer
   sglStart = Timer
   While Timer < sglStart + secs + 1
      intTemp = DoEvents()
   Wend

End Sub


Sub WriteToLog(sTyp As String, strPMGRDBW As String, sComment As String, sFileAndPathName As String)
   'LOGGER
   
   'USE IT LIKE THIS:

   '''LOGGER
   ''Dim strPMGRDBW     As String
   ''Dim LogFileNo      As Integer
   ''Dim sTyp           As String
   ''Dim sComment       As String
   ''Dim sFAPN          As String
   ''sFAPN = "C:\Temp\LOGGY.TX"
   ''sTyp  = "1025 Run OPTIMIZATION II"
   ''sComment = "Optimize Suggestions for Optional Roll Sizes"
   ''strPMGRDBW = ""
   ''strPMGRDBW = DLookup("PM", "tblCurrPM") & " " & DLookup("Grade", "TblPGReqdGrade") & " " & DLookup("BW", "TblPGReqdBW")
   ''Call RHLog(sTyp, strPMGRDBW, sComment, sFAPN)
   
   
   Dim s      As String
   Dim sComma As String
   sComma = ","
   Open sFileAndPathName For Append As #1  'e.g.  "\\s02afs01.na.ipaper.com\lenox-275\CST\Supply Planning\Trim Help RHC\LogFiles\RHLog.CSV" For Append As #1
   'if this above line causes an error  tHEN YOU MIGHT WANT TO do this    Open VBA.Environ("TEMP") & "\RHLog.CSV" For Append As #1
   s = """" & sTyp & """" & sComma & """" & CStr(Now()) & """" & sComma & """" & VBA.Environ("COMPUTERNAME") & """" & sComma & """" & _
              VBA.Environ("USERNAME") & """" & sComma & """" & strPMGRDBW & """" & sComma & """" & CurrentDb.name & """" & sComma & """" & _
              sComment & """" & sComma & """" & "LogFile-ZZZ-" & sFileAndPathName & """"
   Print #1, s
   Close #1
End Sub


Function WriteFile(sPathAndFile As String, sContents As String) As Boolean
'//===============================================================//
'/|   FUNCTION:  WriteFile                                        |/
'/| PARAMETERS:  sPathAndFile, sContents                          |/
'/|    RETURNS:  TRUE ON SUCCESS, FALSE ON FAILURE                |/
'/|    PURPOSE:  Write the string into the file.                  |/
'/|      USAGE:  i = WriteFile("C:\TEMP\FREEED.TXT", str & "#")   |/
'/|         BY:  Sean                                             |/
'/|       DATE:  11/12/2013                                       |/
'/|    HISTORY:                                                   |/
'//===============================================================//
WriteFile = False
    Open sPathAndFile For Output As #1    ' Open to write file.
    Print #1, sContents
    Close #1                              ' Close file.
WriteFile = True
End Function


Private Sub WriteToTextFile(ByVal vsLine As String)
'**********************************************
'Author  :  RKP
'Date/Ver:  07-31-07/V4
'Input   :
'Output  :
'Comments:
'**********************************************
   On Error GoTo Err_Handler

   Print #mnFileHandleOut, vsLine
   
Err_Handler:
   If Err Then
      'ProcessMsg Err.Number, Err.Description, "", ""
      MsgBox Err.Number & " - " & Err.Description
   End If
End Sub


Public Function SAP_GetSession(ByVal sapInstance As String, ByVal reset As Boolean) As Object
'**********************************************
'Author  :  RKP
'Date/Ver:  11-09-12/V01
'Input   :
'Output  :
'Comments:
'**********************************************
    On Error GoTo Err_Handler

    Dim sapGuiAuto      As Object
    Dim sapApp          As Object
    Dim sapConn         As Object
    Dim sapGuiComponent As Object
    Dim sapSessionCount As Integer
    Dim childCtr        As Integer
    Dim err619          As Boolean
    Dim sapSessionFound As Boolean
    
    sapSessionFound = False
    Set SAP_GetSession = Nothing
    
    Set sapGuiAuto = GetObject("SAPGUI")
    If sapGuiAuto Is Nothing Then
        'MsgBox "No active SAP session found." & vbNewLine & "Please log on to: """ & sapInstance & """ and try the operation again.", vbExclamation, Application.ThisWorkbook.Name
    Else
        Set sapApp = sapGuiAuto.GetScriptingEngine
        If sapApp Is Nothing Then
            'MsgBox "No active SAP session found." & vbNewLine & "Please log on to: """ & sapInstance & """ and try the operation again.", vbExclamation, Application.ThisWorkbook.Name
        Else
            sapSessionCount = sapApp.Children.count
            For childCtr = 0 To sapSessionCount - 1
                DoEvents
                Set sapGuiComponent = sapApp.Children.ElementAt(childCtr)
                If sapGuiComponent.Children.count > 0 Then
                    If sapGuiComponent.Children.ElementAt(0).Info.systemName = sapInstance Then
                        Set SAP_GetSession = sapGuiComponent.Children.ElementAt(0)
                        err619 = False
                        Debug.Print SAP_GetSession.findById("wnd[0]/usr/txtRSYST-BNAME").text
                        If sapSessionFound Then
                            Set SAP_GetSession = sapGuiComponent.Children.ElementAt(0)
                            If reset Then
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                                SAP_GetSession.findById("wnd[0]/tbar[0]/btn[3]").press
                            End If
                        Else
                            Set SAP_GetSession = Nothing
                        End If
                    End If
                End If
            Next
            
'            Set sapConn = sapApp.Children(0)
'            If sapConn Is Nothing Then
'                MsgBox "No active SAP session found." & vbNewLine & "Please log on to: """ & sapInstance & """ and try the operation again.", vbExclamation, Application.ThisWorkbook.Name
'            Else
'                Set SAP_GetSession = sapConn.Children(0)
'            End If
        End If
    End If

Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            If Err.Number = 619 Then
                err619 = True
                If err619 Then
                    sapSessionFound = True
                    Resume Next
                End If
            Else
                'ProcessMsg Err.Number, Err.Description, "", ""
                'MsgBox Err.Number & " - " & Err.Description
            End If
        End If
    End If

    Exit Function
    Resume
End Function


Public Sub SAP_RunSAPQueryGeneric( _
    ByRef session As Object, _
    sUserGroup As String, _
    sQueryName As String, _
    sVariant As String, _
    sFilePath As String, _
    Optional sAction As String = "Replace" _
)
'**********************************************
'Author  :  Anjali
'Date/Ver:  02/16/16
'Input   :
'Output  :
'Comments:  02/16/16 Added bew generic function to run SAP queries
'**********************************************

On Error GoTo Err_Handler

'''USAGE EXAMPLES
'''
'''        'Anjali 02/16/16 Edited the code to include new sub - SAP_RunSAPQueryGeneric for customer master
'''        userGroupName = "APPMBMOS"
'''        sapQueryName = "SOP_CUSTMASTER"
'''        FileName = "APPM_KNA1_E01.txt"
'''        VBA.Kill folderPath & "\" & FileName
'''        SAP_RunSAPQueryGeneric session, userGroupName, sapQueryName, "STANDARD", folderPath & "\" & FileName
'''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
'''
'''        'Anjali 02/16/16 Edited the code to include new sub - SAP_RunSAPQueryGeneric for Material master
'''        userGroupName = "APPMBMOS"
'''        sapQueryName = "SOP_MATLMaster"
'''        FileName = "APPM_MARA_E01.txt"
'''        VBA.Kill folderPath & "\" & FileName
'''        SAP_RunSAPQueryGeneric session, userGroupName, sapQueryName, "STANDARD", folderPath & "\" & FileName
'''        If VBA.Dir(folderPath & "\" & FileName) = "" Then msg = msg & vbNewLine & "Error - Did not download: " & folderPath & "\" & FileName
'''
'''



Dim bSAPFileExistsAtBegin As Boolean
Dim sSAPAction As String

sSAPAction = sAction
If VBA.Dir(sFilePath) = "" Then 'File Does not exist
   bSAPFileExistsAtBegin = False
Else
   bSAPFileExistsAtBegin = True                    'File exist
End If


session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
   
session.findById("wnd[0]/tbar[0]/okcd").text = "SQ00"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/mbar/menu[5]/menu[0]").select
session.findById("wnd[1]/usr/radRAD1").select
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[1]/tbar[0]/btn[29]").press
session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btnAPP_WL_SING").press
session.findById("wnd[2]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = sUserGroup 'APPMBMOS
session.findById("wnd[3]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 8
session.findById("wnd[3]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = sQueryName '"SOP_CUSTMASTER"
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").SetFocus
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 14
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/ctxtRS38R-VARIANT").text = sVariant 'STANDARD
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/txt%PATH").text = sFilePath
session.findById("wnd[0]/usr/txt%PATH").SetFocus
session.findById("wnd[0]/usr/txt%PATH").caretPosition = 64
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").Selected = True 'Header
session.findById("wnd[1]/usr/chkRSAQDOWN-COLUMN").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press 'Transfer
If bSAPFileExistsAtBegin = True Then
   If sSAPAction = "Append" Then
      session.findById("wnd[1]/tbar[0]/btn[11]").press 'Append
   Else
      session.findById("wnd[1]/tbar[0]/btn[0]").press 'Replace
   End If
End If
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


Err_Handler:
    mlLastErr = Err.Number
    msLastErr = Err.Description
    'Function1 = mlLastErr
    If Err Then
        If Err.Number = 49 Then 'Bad DLL calling convention
            mlLastErr = 0
            msLastErr = ""
            Resume Next
        Else
            'ProcessMsg Err.Number, Err.Description, "", ""
            MsgBox Err.Number & " - " & Err.Description
        End If
    End If

    Exit Sub
    Resume
End Sub


Public Function UnwindTable(sInputTable As String, sOutputTable As String, iStaticCols As Integer, iUnwindCols As Integer)

Dim sSQL As String
Dim iOutputCounter As Long
Dim iInputCounter As Long
Dim iInputStaticCounter As Long
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset


'sInputTable = "tbl004MasterData020SeasonalityEdit"
'sOutputTable = "tbl004MasterData021SeasonalityWorking"
'iStaticCols = 2
'iUnwindCols = 12

Set rs = Application.CurrentProject.Connection.Execute("select * from " & sInputTable & " where 0=1")
Set rs2 = Application.CurrentProject.Connection.Execute("select * from " & sOutputTable & " where 0=1")


DoCmd.SetWarnings False

iInputCounter = iStaticCols

DoCmd.RunSQL "Delete * from " & sOutputTable

Do Until iInputCounter = iStaticCols + iUnwindCols
    iOutputCounter = 0
    iInputStaticCounter = 0
    
    sSQL = "INSERT INTO " & sOutputTable & "  ( "
    
    'add output table fields to sql
    Do Until iOutputCounter = iStaticCols + 2
        sSQL = sSQL & rs2.Fields(iOutputCounter).name & ", "
        iOutputCounter = iOutputCounter + 1
    Loop

    sSQL = Left(sSQL, Len(sSQL) - 2) & " ) SELECT "
    
    'add input static fields to sql
    Do Until iInputStaticCounter = iStaticCols
        sSQL = sSQL & rs.Fields(iInputStaticCounter).name & ", "
        iInputStaticCounter = iInputStaticCounter + 1
    Loop
    
    sSQL = sSQL & "'" & rs.Fields(iInputCounter).name & "'" & ", [" & rs.Fields(iInputCounter).name & "] FROM " & sInputTable & " WHERE [" & rs.Fields(iInputCounter).name & "] Is Not Null"
    
    '"INSERT INTO " & sOutputTable & "  ( "Plant, Grade, SeasMonth, SeasValue ) SELECT Plant, Grade," & iCounter & ", [" & iCounter & "] FROM " & sInputTable
        
    DoCmd.RunSQL sSQL
    
    
    iInputCounter = iInputCounter + 1
    
Loop

DoCmd.SetWarnings True

End Function

