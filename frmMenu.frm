VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Digital Clock Events Menu"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2850
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdViewEditEvents 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "View/Edit Digital Clock Events"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdViewEditEvents_Click()
frmEvents.Show
Unload Me
End Sub
