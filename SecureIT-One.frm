VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   12615
   ScaleWidth      =   17880
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   16560
      Top             =   1080
   End
   Begin VB.Image Image1 
      Height          =   15360
      Left            =   0
      Picture         =   "SecureIT-One.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19200
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================
' Disable ALT +TAB
' Disable CTRL + ALT + DEL (Cheats by checking if Task manager Running and using Sendkeys to F4 Shut it)
' Disable F4
' TESTED ONLY ON XP PRO ENGLISH
'===========================================================


' COULD BE ADDED TO STARTUP FOLDER
' SIMPLE TIMER PROGRAM COULD LAUNCH THIS EVERY 15 MINUTES OR WHATEVER



'===========================================================
' ALLOW ESCAPE KEY TO EXIT FOR DEBUGGING ONLY
'===========================================================

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'  Select Case KeyCode
'    Case vbKeyEscape
'    Call Cmdunhide
'    Unload Me
'   End Select
'End Sub


Dim hhkLowLevelKybd As Long
Dim Diskeys As Integer
Private Declare Function ShowWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

'===========================================================
' SET TEXT1.TEXT AS PASSWORD - CASE SENSITIVE - BLOCKS CAP LOCK ASWELL - USE SHIFT KEY TO SWITH UPPER/LOWERCASE
'===========================================================
'===========================================================
'===========================================================
Private Sub Text1_Change()
If Text1.Text = "Ian" Then Unload Me
End Sub
'===========================================================
'===========================================================
'===========================================================
'===========================================================






Private Sub Form_Load()
'===========================================================
' GET SCREEN SIZE AND image size ALLOCATE MAXIMUM SCREEN SIZE FOR FORM
'===========================================================
X = Screen.Width
y = Screen.Height
FormMain.Top = 0
FormMain.Left = 0
FormMain.Width = X
FormMain.Height = y
Image1.Width = X
Image1.Height = y
' Position Code input box relative to screen size
Text1.Top = X * 0.34
Text1.Left = y * 0.26
'===========================================================
' HIDE TASKBAR
'===========================================================
Call Cmdhide
'===========================================================
' DISABLE ALT KEYS ETC
'===========================================================
Call Disablealts
End Sub
'===========================================================
' CHECK FOR AN INSTANCE OF TASK MANAGER AND F4 CLOSE IT IF OPEN
'===========================================================
Private Sub Timer1_Timer()
On Error GoTo err_NoTaskmanagerRunning
    AppActivate "Windows Task manager"
    SendKeys "%{F4}"
Exit Sub
err_NoTaskmanagerRunning:
If Err.Number = 5 Then
    Exit Sub
End If
End Sub
Sub TaskBar(blnValue As Boolean)
    Dim lngHandle As Long
    Dim lngStartButton As Long

    lngHandle = FindWindow("Shell_TrayWnd", "")
    
    If blnValue Then
        ShowWindow lngHandle, 5
    Else
        ShowWindow lngHandle, 0
    End If
End Sub
'===========================================================
' HIDE TASKBAR
'===========================================================
Private Sub Cmdhide()
Dim A As Boolean
A = False
TaskBar (A)
End Sub
'===========================================================
' SHOW TASKBAR
'===========================================================
Private Sub Cmdunhide()
Dim A As Boolean
A = True
TaskBar (A)
End Sub
Private Sub Disablealts()
  If Diskeys = 0 Then
    hhkLowLevelKybd = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)
  Else
    UnhookWindowsHookEx hhkLowLevelKybd
    hhkLowLevelKybd = 0
  End If
End Sub
'===========================================================
' BRING BACK TASK BAR ON EXIT
'===========================================================
Private Sub Form_Unload(Cancel As Integer)
Call Cmdunhide
  If hhkLowLevelKybd <> 0 Then UnhookWindowsHookEx hhkLowLevelKybd
End Sub
