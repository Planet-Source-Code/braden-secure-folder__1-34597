VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secure Folder Engine"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCom 
      Height          =   285
      Left            =   120
      LinkItem        =   "txtCom"
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Timer tmrSearch 
      Interval        =   100
      Left            =   2640
      Top             =   120
   End
   Begin VB.Label Label1 
      Caption         =   "Engine Commands:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()


'enable monitoring by default
SearchEnabled = True
'reset firststart
FirstStart = True
'get the settings for the engine
GetSettings


End Sub


Private Sub Form_Unload(Cancel As Integer)


'set status in registry if exiting
SaveSetting "Secure Folder", "Engine", "Status", "NotRun"


End Sub

Private Sub GetSettings()

Dim tempPass
Dim tempFolder


'Get settings from registry

'get the name of the folder to watch for and the encrypted password
tempFolder = GetSetting("Secure Folder", "Settings", "FolderName", "")
tempPass = GetSetting("Secure Folder", "Settings", "Password", "")

'if engine is just starting then set the searchstring and password
If FirstStart = True Then
  SearchString = tempFolder
  EncryptPass = tempPass
  FirstStart = False
Else
  'if not just starting, check if searchstring has been changed by outside influence
  If SearchString <> tempFolder Then
    'if so then save the original searchstring
    SaveSetting "Secure Folder", "Settings", "FolderName", SearchString
  End If
  'check if Password has been changed by outside influence
  If tempPass <> EncryptPass Then
    'if so then save the original password
    SaveSetting "Secure Folder", "Settings", "Password", EncryptPass
  End If
End If


'get aggresive scanning option
AggressiveScan = GetSetting("Secure Folder", "Settings", "AggressiveScan", "False")

'get the hWnd of the contollers txtFolder so we don't close it
hWndExempt = Val(GetSetting("Secure Folder", "Engine", "hWndExempt"))

'set status in registry
If SearchEnabled = True Then SaveSetting "Secure Folder", "Engine", "Status", "RunEnabled"
If SearchEnabled = False Then SaveSetting "Secure Folder", "Engine", "Status", "RunDisabled"


End Sub

Private Sub tmrSearch_Timer()


GetSettings                         'get updated settings right before scan
If SearchEnabled Then
  hWndDesktop = GetDesktopWindow()  'get desktop's hWnd
  DoSearch hWndDesktop              'do the scan starting from the desktop
End If


End Sub

Private Sub txtCom_Change()


'these are the incoming DDE commands from the controller

'change the password
If Left(txtCom.Text, 9) = "PassWord:" Then
  EncryptPass = Mid(txtCom.Text, 10, Len(txtCom.Text))
End If

'change the searchstring
If Left(txtCom.Text, 7) = "Folder:" Then
  SearchString = Mid(txtCom.Text, 8, Len(txtCom.Text))
End If

'shut the engine down
If LCase(txtCom.Text) = "shutdown" Then
  SaveSetting "Secure Folder", "Engine", "Status", "NotRun"
  End
End If

'disable scanning
If LCase(txtCom.Text) = "disable" Then
  SearchEnabled = False
  SaveSetting "Secure Folder", "Engine", "Status", "RunDisabled"
End If

'enable scanning
If LCase(txtCom.Text) = "enable" Then
  SearchEnabled = True
  SaveSetting "Secure Folder", "Engine", "Status", "RunEnabled"
End If

'show the alert form
If LCase(txtCom.Text) = "showalert" Then frmAlert.Show

'show a pop up about messagebox
If LCase(txtCom.Text) = "about" Then MsgBox "Secure Folder Engine v0.98 Beta" & vbCrLf & "Copyright 2002 by Braden Brisbois", vbOKOnly + vbInformation + vbMsgBoxSetForeground, "Security Alert"

'show the engine's main form
If LCase(txtCom.Text) = "showengine" Then
  frmMain.Show
  frmMain.WindowState = 0
  frmMain.Height = 1605
  frmMain.Width = 3360
End If

'hide the engine's main form
If LCase(txtCom.Text) = "hideengine" Then
  frmMain.WindowState = 1
  frmMain.Hide
End If


End Sub



