Attribute VB_Name = "Module2"
Option Explicit

Public EncryptPass As String

'API calls used in the program
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long


Public Function Encrypt(TextToEncrypt As String) As String


'my original one-way encryption algorithym
Dim CharPos As Integer
Dim CurChar, KeyChar As Byte
Dim BitCollector As String

'reset the bit collector
BitCollector = ""

'do a bit-wise XOR character by character using the bitcollector as the key
For CharPos = 1 To Len(TextToEncrypt)
 
 'get the character to encrypt
 CurChar = Asc(Mid(TextToEncrypt, CharPos, 1))
 
 If CharPos > 1 Then
   'get the character before the current character position
   'from the bitcollector to use as the XOR key
   KeyChar = Asc(Mid(BitCollector, CharPos - 1, 1))
   'do the XOR and add the result to the bit collector
   BitCollector = BitCollector + Hex((CurChar Xor KeyChar))
 Else
   'this is the first character so use the last character of the password
   'as the XOR key then do the XOR and add the result to the bit collector
   BitCollector = BitCollector + Hex(CurChar Xor Asc(Mid(TextToEncrypt, Len(TextToEncrypt), 1)))
 End If
Next CharPos

'send the encryption result back as a HEX number
Encrypt = "&H" + BitCollector


End Function


Sub Main()


'if already running then exit
If App.PrevInstance Then End
frmAuthorize.Show


End Sub
