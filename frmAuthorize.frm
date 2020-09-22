VERSION 5.00
Begin VB.Form frmAuthorize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Authorization Required"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ClipControls    =   0   'False
   Icon            =   "frmAuthorize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter Your Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmAuthorize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdExit_Click()


'gave up trying so exit
Unload Me
End


End Sub


Private Sub cmdOK_Click()


Dim TempEncryptPass As String

'encrypt the password attempt
TempEncryptPass = Encrypt(txtPassword.Text)

'compare the two encryptions
If TempEncryptPass = EncryptPass Then
  'they match so let us in
  frmController.Show
  Unload Me
Else
  'they don't match so deny entry
  MsgBox "Incorrect Password" & vbCrLf & "Try again.", vbOKOnly + vbExclamation, "Security Alert"
  txtPassword.Text = ""
End If


End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)


'if the [ESC] key is pressed then exit program
If KeyAscii = 27 Then
   cmdExit_Click
End If


End Sub


Private Sub Form_Load()


'if already running then exit
If App.PrevInstance Then End

'get the encrypted password from the registry
EncryptPass = GetSetting("Secure Folder", "Settings", "Password")

'if there isn't a password then alert and go in
If EncryptPass = "" Then
  MsgBox "No password is currently set." & vbCrLf & "Use the 'Change Password' option in the Controller.", vbOKOnly + vbExclamation, "Security Alert"
  frmController.Show
  Unload Me
End If


End Sub

