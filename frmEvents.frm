VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEvents 
   Caption         =   "Current Digital Clock Events"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   Icon            =   "frmEvents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdDeleteEvent 
      Caption         =   "Delete Event"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddEvent 
      Caption         =   "Add Event"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Width           =   1740
   End
   Begin MSComCtl2.UpDown udTime 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   503
      _Version        =   393216
      BuddyControl    =   "txtEventNumbers"
      BuddyDispid     =   196613
      OrigLeft        =   600
      OrigTop         =   720
      OrigRight       =   795
      OrigBottom      =   975
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtEventNumbers 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2400
      Width           =   420
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   3240
      Top             =   3840
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3120
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   72
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":128C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":18C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":1BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":1EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":220E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":2528
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":2842
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":2B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":2E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":3190
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":34AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":37C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":3ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":3DF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":4112
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":442C
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":4746
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":4A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":4D7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":5094
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":53AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":56C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":59E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":5CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":6016
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":6330
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":664A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":6964
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":6C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":6F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":72B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":75CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":78E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":7C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":7F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":8234
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":854E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":8868
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":8B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":8E9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":91B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":94D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":97EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":9B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":9E1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":A138
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":A452
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":A76C
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":AA86
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":ADA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":B0BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":B3D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":B6EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":BA08
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":BD22
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":C03C
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":C356
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":C670
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":C98A
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":CCA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":CFBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":D2D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":D5F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":D90C
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":DC26
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEvents.frx":DF40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtEvents 
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2760
      Width           =   3735
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEvents.frx":E25A
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblTime 
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   615
   End
End
Attribute VB_Name = "frmEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAddEvent_Click()
If dcNoEvents = True Then
    fRuntimeSetHours -1, 0, 0
    udTime.Max = UBound(dcTimerEventTriggers)
    fShowEvents
    txtEventNumbers = 1
    udTime.Min = 1
    Else
    fRuntimeSetHours -1, 0, 0
    udTime.Max = UBound(dcTimerEventTriggers)
    fShowEvents
End If
End Sub
Function fRuntimeSetHours(Hour As Long, Minute As Long, Second As Long) As Long
Dim TimerTriggers As Long
Dim strHours
Dim strMinutes
Dim strSeconds
Dim strTimedEvent As String
Dim OldTimerEvents()
Dim i
If dcNoEvents = True Then
    TimerTriggers = 1
    dcNoEvents = False
    Else
    TimerTriggers = UBound(dcTimerEventTriggers) + 1
End If
If Hour = -1 Then 'This is to Add a new Slot
    If TimerTriggers > 1 Then
        ReDim OldTimerEvents(1 To TimerTriggers - 1)
        For i = 1 To TimerTriggers - 1
            OldTimerEvents(i) = dcTimerEventTriggers(i)
        Next i
    End If
    ReDim dcTimerEventTriggers(1 To TimerTriggers)
    If TimerTriggers > 1 Then
        For i = 1 To TimerTriggers - 1
             dcTimerEventTriggers(i) = OldTimerEvents(i)
        Next i
    End If
    strTimedEvent = "NEW TIMER EVENT"
    Else
    TimerTriggers = txtEventNumbers
    strHours = "" & Hour
    If Len(strHours) = 1 Then strHours = "0" & strHours
    strMinutes = "" & Minute
    If Len(strMinutes) = 1 Then strMinutes = "0" & strMinutes
    strSeconds = "" & Second
    If Len(strSeconds) = 1 Then strSeconds = "0" & strSeconds
    strTimedEvent = strHours & strMinutes & strSeconds
End If
dcTimerEventTriggers(TimerTriggers) = strTimedEvent
fRuntimeSetHours = TimerTriggers
End Function
Private Sub cmdDeleteEvent_Click()
Dim strMsg As String
Dim strTitle As String
Dim Resp
strMsg = "Are you sure you want to delete Event #" & txtEventNumbers & "?"
strTitle = "Deletion Verification"
Resp = MsgBox(strMsg, vbYesNo + vbSystemModal + vbQuestion, strTitle)
If Resp = vbYes Then
    dcTimerEventTriggers(txtEventNumbers) = "DELETED"
    txtTime = "DELETED"
    fShowEvents
End If
End Sub
Private Sub cmdSubmit_Click()
If fRunTimeTriggerGood(txtTime) = True Then
    fRuntimeSetHours Val(Left(txtTime, 2)), Val(Mid(txtTime, 4, 2)), Val(Right(txtTime, 2))
    fShowEvents
Else
    MsgBox "Invalid Time Format, please use HH:MM:SS.", vbOKOnly + vbCritical, "Invalid Entry"
    txtTime = dcTimerEventTriggers(txtEventNumbers)
End If
End Sub
Private Sub Form_Load()
If dcNoEvents = True Then
    txtEventNumbers = 0
    udTime.Max = 0
    udTime.Min = 0
    txtTime = "No Events"
    Else
    txtEventNumbers = 1
    udTime.Min = 1
    udTime.Max = UBound(dcTimerEventTriggers)
    If Len(dcTimerEventTriggers(1)) <> 6 Then
        txtTime = dcTimerEventTriggers(txtEventNumbers)
        Else
        txtTime = Left(dcTimerEventTriggers(1), 2) & ":" & Mid(dcTimerEventTriggers(1), 3, 2) & ":" & Right(dcTimerEventTriggers(1), 2)
    End If
End If
dcIconCount = 1
fShowEvents
End Sub
Function fRunTimeTriggerGood(strTrigger As String) As Boolean
Dim CurrentTesting As Boolean
Dim strTest As String
Dim strHours As String
Dim strMinutes As String
Dim strSeconds As String
Dim strTemp As String
CurrentTesting = True
strTest = strTrigger
If Len(strTest) <> 8 Then
    CurrentTesting = False
    Else
    If Mid(strTest, 3, 1) <> ":" And Mid(strTest, 6, 1) <> ":" Then CurrentTesting = False
    If Mid(strTest, 3, 1) = "." And Mid(strTest, 6, 1) = "." Then CurrentTesting = True
    ' now test for number integrity
    If CurrentTesting = True Then 'So far so good.
        strHours = Left(strTest, 2)
        strMinutes = Mid(strTest, 4, 2)
        strSeconds = Right(strTest, 2)
        'now check to see if there are any odd characters
        strTemp = "" & Val(strHours)
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        If strTemp <> strHours Then CurrentTesting = False
        strTemp = "" & Val(strMinutes)
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        If strTemp <> strMinutes Then CurrentTesting = False
        strTemp = "" & Val(strSeconds)
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        If strTemp <> strSeconds Then CurrentTesting = False
        If CurrentTesting = True Then 'The numbers are good, now check out the values
            If Val(strHours) < 0 Or Val(strHours) > 24 Then CurrentTesting = False
            If Val(strMinutes) < 0 Or Val(strMinutes) > 59 Then CurrentTesting = False
            If Val(strSeconds) < 0 Or Val(strSeconds) > 59 Then CurrentTesting = False
        End If
    End If
End If
fRunTimeTriggerGood = CurrentTesting
End Function
Private Sub fShowEvents()
txtEvents = ""
Dim strText As String
Dim strTab As String
Dim i
If dcNoEvents = True Then
    txtEvents = "There are no events currently set."
    Else
    strTab = Chr(9)
    strText = "Event Number" & strTab & strTab & "Event Time" & vbCrLf
    For i = 0 To UBound(dcTimerEventTriggers)
        If i > 0 Then
            If Len(dcTimerEventTriggers(i)) <> 6 Then
                strText = "" & i & strTab & strTab & dcTimerEventTriggers(i) & vbCrLf
                Else
                strText = "" & i & strTab & strTab & Left(dcTimerEventTriggers(i), 2) & ":" & Mid(dcTimerEventTriggers(i), 3, 2) & ":" & Right(dcTimerEventTriggers(i), 2) & vbCrLf
            End If
        End If
        txtEvents = txtEvents & strText
    Next i
End If
End Sub
Private Sub Timer1_Timer()
dcIconCount = dcIconCount + 1
If dcIconCount > 72 Then dcIconCount = 1
Me.Icon = ImageList1.ListImages(dcIconCount).Picture
End Sub
Private Sub udTime_Change()
If Len(dcTimerEventTriggers(txtEventNumbers)) <> 6 Then
    txtTime = dcTimerEventTriggers(txtEventNumbers)
    Else
    txtTime = Left(dcTimerEventTriggers(txtEventNumbers), 2) & ":" & Mid(dcTimerEventTriggers(txtEventNumbers), 3, 2) & ":" & Right(dcTimerEventTriggers(txtEventNumbers), 2)
End If
End Sub
