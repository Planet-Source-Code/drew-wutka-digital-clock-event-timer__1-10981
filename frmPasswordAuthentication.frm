VERSION 5.00
Begin VB.Form frmPasswordAuthentication 
   Caption         =   "Password Authentication"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2295
   ControlBox      =   0   'False
   Icon            =   "frmPasswordAuthentication.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   2295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "#"
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter the password required to goto the Events Form."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmPasswordAuthentication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtPassword = dcPasswordString Then
        Unload Me
        frmEvents.Show
        Else
        MsgBox "Invalid Password.", vbOKOnly + vbSystemModal + vbCritical, "Entry Failure"
        Unload Me
    End If
End If
End Sub

