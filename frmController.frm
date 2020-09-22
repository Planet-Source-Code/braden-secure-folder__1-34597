VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmController 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secure Folder Controller"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmController.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "txtCom"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6375
   Begin VB.CommandButton cmdConsole 
      Caption         =   "Show Console"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   3375
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Engine Status"
            TextSave        =   "Engine Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6385
            MinWidth        =   2469
            Text            =   "Getting Engine Status..."
            TextSave        =   "Getting Engine Status..."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "8:56 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameConsole 
      Caption         =   "Engine Console"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdSendCommand 
         Caption         =   "Send Command to Engine"
         Height          =   375
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtCom 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         LinkItem        =   "txtCom"
         TabIndex        =   18
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdEngine 
      Caption         =   "Shut Down Engine"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Frame framePassword 
      Caption         =   "Password"
      Height          =   1815
      Left            =   4080
      TabIndex        =   16
      Top             =   840
      Width           =   2175
      Begin VB.CommandButton cmdChangePassword 
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtPass2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtPass1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdUpdateSettings 
      Caption         =   "Update Settings"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.Timer tmrGetStatus 
      Interval        =   1000
      Left            =   120
      Top             =   5040
   End
   Begin VB.Frame frameSettings 
      Caption         =   "Settings"
      Height          =   1815
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   3855
      Begin VB.TextBox txtFolder 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
      Begin VB.CheckBox chkRunAtStartup 
         Caption         =   "Start engine at bootup"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkAggressiveScan 
         Caption         =   "Aggressive scanning"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox chkDisable 
         Caption         =   "Disable monitoring"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Name of folder to secure:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright 2002 by Braden Brisbois"
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   480
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmController.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Secure Folder v0.98 Beta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private GotExempt As Boolean

Const REG_SZ = 1
Const HKEY_LOCAL_MACHINE = &H80000002
Const REGKEY = "Software\Microsoft\Windows\CurrentVersion\Run"
Const KEY_WRITE = &H20006
Const EXEPATH = "C:\Windows\System\SFEngine.exe"


Private Sub cmdChangePassword_Click()


'check if 2 passwords match
If txtPass1.Text = txtPass2.Text Then
  'check if password is at least 5 characters
  If Len(txtPass1.Text) < 5 Then
    MsgBox "Password must be at least 5 characters long" & vbCrLf & "but no more than 20 characters long", vbOKOnly + vbExclamation, "Security Alert"
  Else
    'encrypt the password
    EncryptPass = Encrypt(txtPass1.Text)
    'if the engine is not runnining then save the encrypted password to the registry
    If cmdEngine.Caption = "Start Engine" Then
      SaveSetting "Secure Folder", "Settings", "Password", EncryptPass
    Else
      'if the engine is running then send it to the engine to be saved
      txtCom.Text = "PassWord:" & EncryptPass
      With txtCom
        .LinkTopic = "SecureFolderEngine|frmMain"
        On Error Resume Next
        .LinkItem = "txtCom"
        .LinkMode = 2
        .LinkPoke
      End With
    End If
    
    MsgBox "Password successfully changed.", vbOKOnly + vbInformation, "Security Alert"
  End If
Else
  'notify that passwords dont match
  MsgBox "Both password fields must match exactly" & vbCrLf & "in order to change the password", vbOKOnly + vbExclamation, "Security Alert"
End If

'clear the password boxes
txtPass1.Text = ""
txtPass2.Text = ""


End Sub


Private Sub cmdConsole_Click()


'show/hide the engine console
If frmController.Height = 4125 Then
  'console is hidden so show it
  frameConsole.Visible = True
  frmController.Height = 4965
  cmdConsole.Caption = "Hide Console"
Else
  'console is visable so hide it
  frameConsole.Visible = False
  frmController.Height = 4125
  cmdConsole.Caption = "Show Console"
End If


End Sub


Private Sub cmdEngine_Click()


'start/shutdown the engine
On Error GoTo err:

If cmdEngine.Caption = "Start Engine" Then
  'engine is not running so start it
  Shell "C:\Windows\System\SFEngine.exe"
Else
  'engine is running so shut it down
  txtCom.Text = "ShutDown"
  'send DDE message to engine to shutdown
  With txtCom
    .LinkTopic = "SecureFolderEngine|frmMain"
    On Error Resume Next
    .LinkItem = "txtCom"
    .LinkMode = 2
    .LinkPoke
  End With
End If

'clear button caption
cmdEngine.Caption = ""

Exit Sub


'the engine .EXE couldn't be found, show error message to fix it
err:
If Error = "File not found" Then
  MsgBox "Can not find Secure Folder Engine." & vbCrLf & _
         "Folder protection can not be performed." & vbCrLf & vbCrLf & _
         "Place the 'SFENGINE.EXE' file in the 'C:\WINDOWS\SYSTEM\' folder" & vbCrLf & _
         "and try starting it again.", vbOKOnly + vbCritical, "Security Alert"
End If
Resume Next


End Sub


Private Sub cmdExit_Click()

Unload Me

End Sub


Private Sub cmdUpdateSettings_Click()

UpdateSettings

End Sub


Private Sub cmdSendCommand_Click()


'send a manual DDE command to the engine
With txtCom
  .LinkTopic = "SecureFolderEngine|frmMain"
  On Error Resume Next
  .LinkItem = "txtCom"
  .LinkMode = 2
  .LinkPoke
End With


End Sub

Private Sub Form_Load()


'if already running then exit
If App.PrevInstance Then End


'get window position from registry
frmController.Top = GetSetting("Secure Folder", "Window Position", "Top", 0)
frmController.Left = GetSetting("Secure Folder", "Window Position", "Left", 0)


'init controller
GotExempt = False
frmController.Show


'clear out engine status in registry. if engine is running it'll reset it
SaveSetting "Secure Folder", "Engine", "Status", "NotRun"


'set the txtFolder's hWnd in registry so the engine won't close it
SaveSetting "Secure Folder", "Engine", "hWndExempt", Str(txtFolder.hWnd)


'get aggressive scan option from registry and set the checkbox
If GetSetting("Secure Folder", "Settings", "AggressiveScan", "False") = "True" Then
  chkAggressiveScan.Value = vbChecked
Else
  chkAggressiveScan.Value = vbUnchecked
End If


'get start engine at bootup option and set the checkbox
If GetSetting("Secure Folder", "Settings", "StartAtBoot", "True") = "True" Then
  chkRunAtStartup.Value = vbChecked
Else
  chkRunAtStartup.Value = vbUnchecked
End If


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


'Exiting the program! save the settings!
UpdateSettings
'save the window position
SaveSetting "Secure Folder", "Window Position", "Top", frmController.Top
SaveSetting "Secure Folder", "Window Position", "Left", frmController.Left
'clear out the exempted txtFolder's hWnd
SaveSetting "Secure Folder", "Engine", "hWndExempt", "0"


End Sub


Private Sub tmrGetStatus_Timer()


Dim Status As String


'get engine status from registry
Status = GetSetting("Secure Folder", "Engine", "Status")

'then display the status in the status bar
If Status = "NotRun" Then
  'engine is not running!
  StatusBar1.Panels.Item(2).Text = "Not Loaded!"
  cmdEngine.Caption = "Start Engine"
End If

If Status = "RunDisabled" Then
  'engine is running but not monitoring
  StatusBar1.Panels.Item(2).Text = "Running but NOT Monitoring"
  cmdEngine.Caption = "Shut Down Engine"
End If

If Status = "RunEnabled" Then
  'engine is running and monitoring
  StatusBar1.Panels.Item(2).Text = "Running and Monitoring..."
  cmdEngine.Caption = "Shut Down Engine"
End If


'check to see if txtFolder has been retrieved from registry yet.
'must do it this way to give the engine time to receive txtFolder's hWnd
If GotExempt = False Then
 txtFolder.Text = GetSetting("Secure Folder", "Settings", "FolderName", "No Folder Specified")
 GotExempt = True
End If


End Sub


Private Sub UpdateSettings()


'update the settings of the program in the registry
Dim hWndReg As Long
Dim temp As Long
Dim a As Integer


'check the length of folder name to protect and save it if > 5 characters
If Len(txtFolder.Text) < 5 Then
  'name is less than 5, notify and exit sub without saving anything
  MsgBox "The name of the folder to protect must" & vbCrLf & "be at least 5 characters long." & vbCrLf & vbCrLf & "Settings were NOT updated.", vbOKOnly + vbExclamation, "Security Alert"
  Exit Sub
Else
  'if the engine is running send the folder name to the engine to be saved
  If cmdEngine.Caption <> "Start Engine" Then
    txtCom.Text = "Folder:" & txtFolder.Text
    With txtCom
      .LinkTopic = "SecureFolderEngine|frmMain"
      On Error Resume Next
      .LinkItem = "txtCom"
      .LinkMode = 2
      .LinkPoke
    End With
  Else
    'if it's not running then save it to the registry
    SaveSetting "Secure Folder", "Settings", "FolderName", txtFolder.Text
  End If
End If


'get aggressive scanning enabled/disabled and set option in registry
If chkAggressiveScan.Value = vbChecked Then
  SaveSetting "Secure Folder", "Settings", "AggressiveScan", "True"
Else
  SaveSetting "Secure Folder", "Settings", "AggressiveScan", "False"
End If


'update start at boot option
If chkRunAtStartup.Value = vbChecked Then
  'save option to registry
  SaveSetting "Secure Folder", "Settings", "StartAtBoot", "True"
  'write run entry to registry
  temp = RegOpenKeyEx(HKEY_LOCAL_MACHINE, REGKEY, 0, KEY_WRITE, hWndReg)
  RegSetValueEx hWndReg, "SecureFolder", 0, REG_SZ, ByVal EXEPATH, Len(EXEPATH)
Else
  'save option to registry
  SaveSetting "Secure Folder", "Settings", "StartAtBoot", "False"
  'delete run entry from registry
  temp = RegOpenKeyEx(HKEY_LOCAL_MACHINE, REGKEY, 0, KEY_WRITE, hWndReg)
  RegDeleteValue hWndReg, "SecureFolder"
End If


'send DDE message to engine for scanning enabled/disabled
'figure which message to send and put it in txtCom.text
If chkDisable.Value = vbChecked Then
  txtCom.Text = "Disable"
Else
  txtCom.Text = "Enable"
End If
'send the message
With txtCom
  .LinkTopic = "SecureFolderEngine|frmMain"
  On Error Resume Next
  .LinkItem = "txtCom"
  .LinkMode = 2
  .LinkPoke
End With


End Sub

Private Sub txtFolder_GotFocus()


'highlight all the text in txtFolder
txtFolder.SelStart = 0
txtFolder.SelLength = Len(txtFolder.Text)


End Sub
