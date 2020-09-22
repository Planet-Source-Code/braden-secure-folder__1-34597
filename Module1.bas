Attribute VB_Name = "Module1"
Option Explicit

Public AggressiveScan As Boolean
Public SearchEnabled As Boolean
Public SearchString As String
Public hWndDesktop As Long
Public hWndExempt As Long
Public FirstStart As Boolean
Public EncryptPass As String


'constants used in the engine
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_CLOSE = &H10


'API calls used in the engine
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function RegisterServiceProcess Lib "Kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long


Public Sub MakeWindowAlwaysTop(hwnd As Long)

'make a form always on top
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE

End Sub


Public Sub MakeWindowNotTop(hwnd As Long)
    
'un-"ontop" a form
SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE

End Sub


Public Sub AltCtrlDel_Show()
  
'show the engine in the CTRL-ALT-DEL list
Call RegisterServiceProcess(0, 0)

End Sub


Public Sub AltCtrlDel_Hide()
    
'hdie the engine from the CTRL-ALT-DEL list
Call RegisterServiceProcess(0, 1)

End Sub


Public Sub DoSearch(hWndParent As Long)

Dim hWndChild As Long
Dim length As Long
Dim result As Long
Dim strtmp As String

'Get the first child of hWndParent
hWndChild = GetWindow(hWndParent, GW_CHILD Or GW_HWNDFIRST)

Do While hWndChild <> 0
  
  'hWndChild contains a child of hWndParent
  'get window text of hWndChild
  length = SendMessage(hWndChild, WM_GETTEXTLENGTH, ByVal 0, ByVal 0) + 1
  'error trap in case length of text is HUGE and runs out of string space
  If length > 1024 Then length = 1024
  strtmp = Space(length)
  result = SendMessage(hWndChild, WM_GETTEXT, ByVal length, ByVal strtmp)
  
  'check to see if search string is in object
  If InStr(1, strtmp, SearchString) > 0 And hWndChild <> hWndExempt Then
    'text was found, show 'Access Denied'
    frmAlert.Show
    'close the object
    result = PostMessage(hWndChild, WM_CLOSE, 0, 0)
    'if object's parent is not the desktop then close it also
    If hWndParent <> hWndDesktop Then
      result = PostMessage(hWndParent, WM_CLOSE, 0, 0)
    End If
  End If
  
  'Now get any children for hWndChild if AggressiveScan is true
  If AggressiveScan Then DoSearch hWndChild
  
  'move on to the next window
  hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
Loop


End Sub


Sub Main()
    
    If App.PrevInstance Then End   'if already running then end
    AltCtrlDel_Hide                'take application out of the Ctrl-Alt-Del list
    frmMain.Hide                   'load form but don't show it
    
End Sub

