VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl DigitalClock 
   Alignable       =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   ScaleHeight     =   2145
   ScaleWidth      =   8025
   ToolboxBitmap   =   "DigitalClock.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   1320
   End
   Begin MSComctlLib.ImageList ilnumbers 
      Index           =   0
      Left            =   1320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   200
      ImageHeight     =   305
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":0F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":2462
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":38FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":4B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":5FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":7762
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":868A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":A0C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":B826
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":CFE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":D682
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":DB56
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":E00A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilnumbers 
      Index           =   1
      Left            =   720
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   200
      ImageHeight     =   305
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":F23C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":FCE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":10D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":11DF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":12C84
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":13D10
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":14FA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":15C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":170B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":1830C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":195C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":19C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":1A084
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":1A494
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilnumbers 
      Index           =   2
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   200
      ImageHeight     =   305
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":1B336
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":1C10E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":1D966
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":1F176
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":20632
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":21E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":239CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":24B1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":269FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":28552
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":2A10A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":2A7AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":2AD16
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DigitalClock.ctx":2B266
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgAmPm 
      Height          =   255
      Left            =   7440
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape Colin2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   5220
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Colin2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   5220
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Colin1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   2580
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Colin1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   2580
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Image imgNumber 
      Height          =   1695
      Index           =   5
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image imgNumber 
      Height          =   1695
      Index           =   4
      Left            =   5520
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image imgNumber 
      Height          =   1695
      Index           =   3
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image imgNumber 
      Height          =   1695
      Index           =   2
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image imgNumber 
      Height          =   1695
      Index           =   1
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image imgNumber 
      Height          =   1695
      Index           =   0
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "DigitalClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Const dcGreen = &HFF00&
Const dcRed = &HFF&
Const dcBlue = &HFF0000
Dim dcColorScheme As Long
Dim dcPrePropsHaveBeenSet As Boolean
'Default Property Values:
Const m_def_LockColors = False
Const m_def_LockTimeFormat = False
Const m_def_LockEventDisplay = False
Const m_def_LockMenu = False
Const m_def_MenuPassword = ""
Const m_def_PreEvent1 = ""
Const m_def_PreEvent2 = ""
Const m_def_PreEvent3 = ""
Const m_def_PreEvent4 = ""
Const m_def_PreEvent5 = ""
Const m_def_PreEvent6 = ""
Const m_def_PreEvent7 = ""
Const m_def_PreEvent8 = ""
Const m_def_PreEvent9 = ""
Const m_def_PreEvent10 = ""
Const m_def_SetTimeHours = 0
Const m_def_SetTimeMinutes = 0
Const m_def_SetTimeSeconds = 0
Const m_def_MilitaryTime = True
Const m_def_NumberColor = 0
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONUP = &H205
'Property Variables:
Dim m_LockColors As Boolean
Dim m_LockTimeFormat As Boolean
Dim m_LockEventDisplay As Boolean
Dim m_LockMenu As Boolean
Dim m_MenuPassword As String
Dim m_PreEvent1 As String
Dim m_PreEvent2 As String
Dim m_PreEvent3 As String
Dim m_PreEvent4 As String
Dim m_PreEvent5 As String
Dim m_PreEvent6 As String
Dim m_PreEvent7 As String
Dim m_PreEvent8 As String
Dim m_PreEvent9 As String
Dim m_PreEvent10 As String
Dim m_SetTimeHours As Variant
Dim m_SetTimeMinutes As Variant
Dim m_SetTimeSeconds As Variant
Dim m_MilitaryTime As Boolean
Dim m_NumberColor As Long
Dim dcPreviousTime As String
Dim dcPreviousAmPm As Long
Dim dcEventFireTime As String
Dim dcEventFired As Boolean
'Event Declarations:
Event TimeReached(SetTimeIndex As Integer)
''Event TimeReached(SetTimeIndex As Integer) 'MappingInfo=Label1(0),Label1,0,Change
''Event TimeReached(SetTimeIndex As Long)
'Function fSetNumbers(NumToSet As Long, GroupToSet As Long)
'Dim PicNum
'If NumToSet > 0 Then PicNum = NumToSet
'If NumToSet = 0 Then PicNum = 10
'If NumToSet = -1 Then PicNum = 11
'imgNumber(GroupToSet).Picture = ilnumbers(dcColorScheme).ListImages.Item(PicNum).Picture
'End Function
Function fRunTREvent(EventIndex As Integer)
Dim dtStartEvent As Date
Dim dtEndEvent As Date
Dim EventNumber As Integer
Dim i As Long
Timer1.Interval = 0
DoEvents
dtStartEvent = Now()
dcPreviousTime = "Event"
fSetJobDisplay EventIndex
DoEvents
RaiseEvent TimeReached(EventIndex)
DoEvents
dtEndEvent = Now()
EventNumber = EventIndex
dcPreviousTime = "EventsFinished"
Timer1.Interval = 100
End Function
Function fSetJobDisplay(EventNumber As Integer)
'now set the clock to show which event is running.
Dim strEventNumber As String
Dim intImageNumber As Long
Dim i
imgNumber(0).Picture = ilnumbers(dcColorScheme).ListImages.Item(14).Picture
imgNumber(1).Picture = ilnumbers(dcColorScheme).ListImages.Item(10).Picture
imgNumber(2).Picture = ilnumbers(dcColorScheme).ListImages.Item(8).Picture
strEventNumber = "" & EventNumber
Do Until Len(strEventNumber) = 3
    strEventNumber = "0" & strEventNumber
Loop
For i = 1 To 3
    intImageNumber = Val(Mid(strEventNumber, i, 1))
    If intImageNumber = 0 Then intImageNumber = 10
    imgNumber(i + 2).Picture = ilnumbers(dcColorScheme).ListImages.Item(intImageNumber).Picture
Next i
End Function
Function fSetTime()
Dim strTrigger As String
Dim strHours As String
Dim strMinutes As String
Dim strSeconds As String
If dcPrePropsHaveBeenSet = False Then
    Dim l
    'Set up the Design Time Triggers
    For l = 1 To 10
        strTrigger = fGetPresetTriggers(l)
        If fDesignTriggerGood(strTrigger) Then
            strHours = Left(strTrigger, 2)
            strMinutes = Mid(strTrigger, 4, 2)
            strSeconds = Right(strTrigger, 2)
            SetTime Val(strHours), Val(strMinutes), Val(strSeconds)
            Else
            GoTo DoneWithPresets
        End If
    Next l
DoneWithPresets:
    dcPrePropsHaveBeenSet = True
    Timer1.Interval = 100
End If
If dcPreviousTime <> "Event" Then
    Dim CurTime As Date
    Dim TempNum As Long
    Dim TempNow As String
    CurTime = Now()
    strHours = "" & Format(CurTime, "hh")
    strMinutes = "" & Format(CurTime, "nn")
    strSeconds = "" & Format(CurTime, "ss")
    If strHours = "24" Then strHours = "00"
    TempNow = strHours & strMinutes & strSeconds
    If dcPreviousTime <> TempNow Then
        If dcNoEvents = False Then
            Dim i As Integer
            Dim se As Long
            se = 0
            If dcEventFired = True Then
                dcEventFired = False
                Dim TempRepeatLoop As String
                TempRepeatLoop = dcEventFireTime
                If Val(TempNow) < Val(TempRepeatLoop) Then
                    For i = 1 To UBound(dcTimerEventTriggers)
                        If Val(dcTimerEventTriggers(i)) <= 250000 And Val(dcTimerEventTriggers(i)) > Val(TempRepeatLoop) Then
                            dcEventFired = True
                            dcEventFireTime = TempNow
                            fRunTREvent i
                            'i = UBound(dcTimerEventTriggers)
                            'Exit Function
                        End If
                    Next i
                    For i = 1 To UBound(dcTimerEventTriggers)
                        If Val(dcTimerEventTriggers(i)) <= Val(TempNow) And Val(dcTimerEventTriggers(i)) >= 0 Then
                            dcEventFired = True
                            dcEventFireTime = TempNow
                            fRunTREvent i
                            'i = UBound(dcTimerEventTriggers)
                            'Exit Function
                        End If
                    Next i
                    Else
                    For i = 1 To UBound(dcTimerEventTriggers)
                        If Val(dcTimerEventTriggers(i)) <= Val(TempNow) And Val(dcTimerEventTriggers(i)) > Val(TempRepeatLoop) Then
                            dcEventFired = True
                            dcEventFireTime = TempNow
                            fRunTREvent i
                            'i = UBound(dcTimerEventTriggers)
                            'Exit Function
                        End If
                    Next i
                End If
            Else
                For i = 1 To UBound(dcTimerEventTriggers)
                    If dcTimerEventTriggers(i) = TempNow Then
                        dcEventFired = True
                        dcEventFireTime = TempNow
                        fRunTREvent i
                        'i = UBound(dcTimerEventTriggers)
                        'Exit Function
                    End If
                Next i
            End If
        End If
        If Me.MilitaryTime = True Then
            If dcPreviousAmPm <> 0 Then
                imgAmPm.Picture = ilnumbers(dcColorScheme).ListImages.Item(11).Picture
                dcPreviousAmPm = 0
            End If
            Else
            Dim AmPm As Long
           If Val(strHours) > 11 Then
                AmPm = 13
                If Val(strHours) > 12 Then
                    strHours = "" & Val(strHours) - 12
                    If Len(strHours) = 1 Then strHours = "0" & strHours
                End If
                Else
                AmPm = 12
                If Val(strHours) = "00" Then strHours = "12"
            End If
            If dcPreviousAmPm <> AmPm Then
               imgAmPm.Picture = ilnumbers(dcColorScheme).ListImages.Item(AmPm).Picture
                dcPreviousAmPm = AmPm
            End If
        End If
        If Mid(dcPreviousTime, 1, 1) <> Left(strHours, 1) Then
            If Left(strHours, 1) = "0" Then
                TempNum = 11
                If Me.MilitaryTime = True Then TempNum = 10
               Else
                TempNum = Val(Left(strHours, 1))
            End If
            fSetNumbers TempNum, 0
        End If
        If Mid(dcPreviousTime, 2, 1) <> Right(strHours, 1) Then
           If Right(strHours, 1) = "0" Then
                TempNum = 10
                Else
                TempNum = Val(Right(strHours, 1))
            End If
            fSetNumbers TempNum, 1
        End If
        If Mid(dcPreviousTime, 3, 1) <> Left(strMinutes, 1) Then
            If Left(strMinutes, 1) = "0" Then
                TempNum = 10
                Else
                TempNum = Val(Left(strMinutes, 1))
            End If
        fSetNumbers TempNum, 2
        End If
        If Mid(dcPreviousTime, 4, 1) <> Right(strMinutes, 1) Then
            If Right(strMinutes, 1) = "0" Then
                TempNum = 10
                Else
                TempNum = Val(Right(strMinutes, 1))
            End If
            fSetNumbers TempNum, 3
        End If
        If Mid(dcPreviousTime, 5, 1) <> Left(strSeconds, 1) Then
            If Left(strSeconds, 1) = "0" Then
                TempNum = 10
                Else
                TempNum = Val(Left(strSeconds, 1))
            End If
            fSetNumbers TempNum, 4
        End If
        If Mid(dcPreviousTime, 6, 1) <> Right(strSeconds, 1) Then
            If Right(strSeconds, 1) = "0" Then
                TempNum = 10
                Else
                TempNum = Val(Right(strSeconds, 1))
            End If
            fSetNumbers TempNum, 5
        End If
        dcPreviousTime = strHours & strMinutes & strSeconds
    End If
End If
End Function

Private Sub imgNumber_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TestColor
Select Case Button
    Case 1 '514 Run Left Click Routine(ColorChange)
    If Me.LockColors = False Then
        TestColor = Me.NumberColor
        TestColor = TestColor + 1
        If TestColor = 3 Then TestColor = 0
        Me.NumberColor = TestColor
    End If
    Case 2 '517 Run Right Click(Military/12hour time switch)
    If Me.LockTimeFormat = False Then MilitaryTime = Not MilitaryTime
End Select
End Sub
Private Sub Timer1_Timer()
fSetTime
End Sub
Private Sub UserControl_Initialize()
dcNoEvents = True
dcEventFired = False
'dcPrePropsHaveBeenSet = True
fSetTime
dcPrePropsHaveBeenSet = False
End Sub
Function fGetPresetTriggers(TrigNum) As String
If TrigNum = 1 Then fGetPresetTriggers = PreEvent1
If TrigNum = 2 Then fGetPresetTriggers = PreEvent2
If TrigNum = 3 Then fGetPresetTriggers = PreEvent3
If TrigNum = 4 Then fGetPresetTriggers = PreEvent4
If TrigNum = 5 Then fGetPresetTriggers = PreEvent5
If TrigNum = 6 Then fGetPresetTriggers = PreEvent6
If TrigNum = 7 Then fGetPresetTriggers = PreEvent7
If TrigNum = 8 Then fGetPresetTriggers = PreEvent8
If TrigNum = 9 Then fGetPresetTriggers = PreEvent9
If TrigNum = 10 Then fGetPresetTriggers = PreEvent10
End Function
Function fDesignTriggerGood(strTrigger As String) As Boolean
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
fDesignTriggerGood = CurrentTesting
End Function
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TestColor
Select Case Button
    Case 1 '514 Run Left Click Routine(Display Events)
    If Me.LockEventDisplay = False Then
        'This will display the Current 'Triggers'
        If dcNoEvents = True Then Exit Sub
        Dim strMsg As String
        Dim strTitle As String
        Dim Resp
        Dim i
        Dim strTime As String
        Dim NumOfTriggers As Long
        Dim TriggerMultiplier As Long
        Dim strTab As String
        Dim NumToShow As Long
        Dim KeepGoing As Boolean
        Dim CurrentTrigger As Long
        strTab = Chr(9)
        TriggerMultiplier = 0
        strTitle = "Digital Clock Trigger Times"
        NumOfTriggers = UBound(dcTimerEventTriggers)
        If NumOfTriggers < 11 Then 'Just run this once
            strMsg = "The  Triggers  currently  stored  within  the " & vbCrLf
            strMsg = strMsg & "Digital Clock  are  listed  below.  They  are " & vbCrLf
            strMsg = strMsg & "listed with their Trigger number and Trigger" & vbCrLf
            strMsg = strMsg & "Time (In Military Time)." & vbCrLf & vbCrLf
            For i = 1 To NumOfTriggers
                strTime = Left(dcTimerEventTriggers(i), 2) & ":" & Mid(dcTimerEventTriggers(i), 3, 2) & ":" & Right(dcTimerEventTriggers(i), 2)
                strMsg = strMsg & strTab & i & strTab & strTime & vbCrLf
            Next i
            MsgBox strMsg, vbOKOnly + vbInformation, strTitle
        Else
            KeepGoing = True
            Do Until KeepGoing = False
                NumToShow = NumOfTriggers
                If NumToShow > 10 Then NumToShow = 10
                strMsg = "The  Triggers  currently  stored  within  the " & vbCrLf
                strMsg = strMsg & "Digital Clock  are  listed  below.  They  are " & vbCrLf
                strMsg = strMsg & "listed with their Trigger number and Trigger" & vbCrLf
                strMsg = strMsg & "Time (In Military Time)." & vbCrLf & vbCrLf
                For i = 1 To NumToShow
                    CurrentTrigger = i + (10 * TriggerMultiplier)
                    strTime = Left(dcTimerEventTriggers(CurrentTrigger), 2) & ":" & Mid(dcTimerEventTriggers(CurrentTrigger), 3, 2) & ":" & Right(dcTimerEventTriggers(CurrentTrigger), 2)
                    strMsg = strMsg & strTab & CurrentTrigger & strTab & strTime & vbCrLf
                Next i
                NumOfTriggers = NumOfTriggers - 10
                TriggerMultiplier = TriggerMultiplier + 1
                If NumOfTriggers > 0 Then
                    strMsg = strMsg & vbCrLf & "Press 'Yes' to see more Triggers." & vbCrLf & "Press 'No' to close this window." & vbCrLf
                    Resp = MsgBox(strMsg, vbYesNo + vbSystemModal + vbInformation, strTitle)
                    If Resp = vbYes Then KeepGoing = True
                    If Resp = vbNo Then KeepGoing = False
                    Else
                    KeepGoing = False
                    MsgBox strMsg, vbOKOnly + vbSystemModal + vbInformation, strTitle
                End If
            Loop
        End If
    End If
    Case 2 '517 Run Right Click(Display Menu)
    If Me.LockMenu = False Then
        If Me.MenuPassword <> "" Then 'A password has been set.
            frmPasswordAuthentication.Show
            Else
            frmEvents.Show
        End If
    End If
End Select
End Sub

Private Sub UserControl_Resize()
Dim ctrlHieght As Double
Dim ctrlWidth As Double
ctrlHieght = UserControl.ScaleHeight
ctrlWidth = UserControl.ScaleWidth
Dim i
For i = 0 To 5
    imgNumber(i).Height = 1695 * (ctrlHieght / 2145)
    imgNumber(i).Width = 1095 * (ctrlWidth / 8025)
    imgNumber(i).Top = 240 * (ctrlHieght / 2145)
Next i
imgNumber(0).Left = 240 * (ctrlWidth / 8025)
imgNumber(1).Left = 1440 * (ctrlWidth / 8025)
imgNumber(2).Left = 2880 * (ctrlWidth / 8025)
imgNumber(3).Left = 4080 * (ctrlWidth / 8025)
imgNumber(4).Left = 5520 * (ctrlWidth / 8025)
imgNumber(5).Left = 6720 * (ctrlWidth / 8025)
Dim TempH
For i = 0 To 1
    Colin1(i).Height = 375 * (ctrlHieght / 2145)
    Colin1(i).Width = 255 * (ctrlWidth / 8025)
    Colin2(i).Height = 375 * (ctrlHieght / 2145)
    Colin2(i).Width = 255 * (ctrlWidth / 8025)
    If i = 0 Then TempH = 600
    If i = 1 Then TempH = 1080
    Colin1(i).Top = TempH * (ctrlHieght / 2145)
    Colin2(i).Top = TempH * (ctrlHieght / 2145)
    Colin1(i).Left = 2580 * (ctrlWidth / 8250)
    Colin2(i).Left = 5220 * (ctrlWidth / 8250)
Next i
imgAmPm.Height = 255 * (ctrlHieght / 2145)
imgAmPm.Width = 615 * (ctrlWidth / 8025)
imgAmPm.Top = ctrlHieght - imgAmPm.Height
imgAmPm.Left = ctrlWidth - imgAmPm.Width
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get NumberColor() As Long
Attribute NumberColor.VB_Description = "Set to 0 for Red, 1 for Green, and 2 for Blue."
    NumberColor = m_NumberColor
End Property

Public Property Let NumberColor(ByVal New_NumberColor As Long)
    dcPreviousTime = ""
    m_NumberColor = New_NumberColor
    PropertyChanged "NumberColor"
    dcColorScheme = New_NumberColor
    Dim NewColor
    If New_NumberColor = 0 Then NewColor = dcRed
    If New_NumberColor = 1 Then NewColor = dcBlue
    If New_NumberColor = 2 Then NewColor = dcGreen
    Dim i
    For i = 0 To 1
        Colin1(i).FillColor = NewColor
        Colin1(i).BorderColor = NewColor
        Colin2(i).FillColor = NewColor
        Colin2(i).BorderColor = NewColor
    Next i
    fSetTime
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function SetTime(Hour As Long, Minute As Long, Second As Long) As Long
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
strHours = "" & Hour
If Len(strHours) = 1 Then strHours = "0" & strHours
strMinutes = "" & Minute
If Len(strMinutes) = 1 Then strMinutes = "0" & strMinutes
strSeconds = "" & Second
If Len(strSeconds) = 1 Then strSeconds = "0" & strSeconds
strTimedEvent = strHours & strMinutes & strSeconds
dcTimerEventTriggers(TimerTriggers) = strTimedEvent
SetTime = TimerTriggers
End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_NumberColor = m_def_NumberColor
    m_MilitaryTime = m_def_MilitaryTime
    m_SetTimeHours = m_def_SetTimeHours
    m_SetTimeMinutes = m_def_SetTimeMinutes
    m_SetTimeSeconds = m_def_SetTimeSeconds
    m_PreEvent1 = m_def_PreEvent1
    m_PreEvent2 = m_def_PreEvent2
    m_PreEvent3 = m_def_PreEvent3
    m_PreEvent4 = m_def_PreEvent4
    m_PreEvent5 = m_def_PreEvent5
    m_PreEvent6 = m_def_PreEvent6
    m_PreEvent7 = m_def_PreEvent7
    m_PreEvent8 = m_def_PreEvent8
    m_PreEvent9 = m_def_PreEvent9
    m_PreEvent10 = m_def_PreEvent10
    m_LockColors = m_def_LockColors
    m_LockTimeFormat = m_def_LockTimeFormat
    m_LockEventDisplay = m_def_LockEventDisplay
    m_LockMenu = m_def_LockMenu
    m_MenuPassword = m_def_MenuPassword
dcColorScheme = m_def_NumberColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_NumberColor = PropBag.ReadProperty("NumberColor", m_def_NumberColor)
    m_MilitaryTime = PropBag.ReadProperty("MilitaryTime", m_def_MilitaryTime)
    m_SetTimeHours = PropBag.ReadProperty("SetTimeHours", m_def_SetTimeHours)
    m_SetTimeMinutes = PropBag.ReadProperty("SetTimeMinutes", m_def_SetTimeMinutes)
    m_SetTimeSeconds = PropBag.ReadProperty("SetTimeSeconds", m_def_SetTimeSeconds)
    m_PreEvent1 = PropBag.ReadProperty("PreEvent1", m_def_PreEvent1)
    m_PreEvent2 = PropBag.ReadProperty("PreEvent2", m_def_PreEvent2)
    m_PreEvent3 = PropBag.ReadProperty("PreEvent3", m_def_PreEvent3)
    m_PreEvent4 = PropBag.ReadProperty("PreEvent4", m_def_PreEvent4)
    m_PreEvent5 = PropBag.ReadProperty("PreEvent5", m_def_PreEvent5)
    m_PreEvent6 = PropBag.ReadProperty("PreEvent6", m_def_PreEvent6)
    m_PreEvent7 = PropBag.ReadProperty("PreEvent7", m_def_PreEvent7)
    m_PreEvent8 = PropBag.ReadProperty("PreEvent8", m_def_PreEvent8)
    m_PreEvent9 = PropBag.ReadProperty("PreEvent9", m_def_PreEvent9)
    m_PreEvent10 = PropBag.ReadProperty("PreEvent10", m_def_PreEvent10)
    m_LockColors = PropBag.ReadProperty("LockColors", m_def_LockColors)
    m_LockTimeFormat = PropBag.ReadProperty("LockTimeFormat", m_def_LockTimeFormat)
    m_LockEventDisplay = PropBag.ReadProperty("LockEventDisplay", m_def_LockEventDisplay)
    m_LockMenu = PropBag.ReadProperty("LockMenu", m_def_LockMenu)
    m_MenuPassword = PropBag.ReadProperty("MenuPassword", m_def_MenuPassword)
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("NumberColor", m_NumberColor, m_def_NumberColor)
    Call PropBag.WriteProperty("MilitaryTime", m_MilitaryTime, m_def_MilitaryTime)
    Call PropBag.WriteProperty("SetTimeHours", m_SetTimeHours, m_def_SetTimeHours)
    Call PropBag.WriteProperty("SetTimeMinutes", m_SetTimeMinutes, m_def_SetTimeMinutes)
    Call PropBag.WriteProperty("SetTimeSeconds", m_SetTimeSeconds, m_def_SetTimeSeconds)
    Call PropBag.WriteProperty("PreEvent1", m_PreEvent1, m_def_PreEvent1)
    Call PropBag.WriteProperty("PreEvent2", m_PreEvent2, m_def_PreEvent2)
    Call PropBag.WriteProperty("PreEvent3", m_PreEvent3, m_def_PreEvent3)
    Call PropBag.WriteProperty("PreEvent4", m_PreEvent4, m_def_PreEvent4)
    Call PropBag.WriteProperty("PreEvent5", m_PreEvent5, m_def_PreEvent5)
    Call PropBag.WriteProperty("PreEvent6", m_PreEvent6, m_def_PreEvent6)
    Call PropBag.WriteProperty("PreEvent7", m_PreEvent7, m_def_PreEvent7)
    Call PropBag.WriteProperty("PreEvent8", m_PreEvent8, m_def_PreEvent8)
    Call PropBag.WriteProperty("PreEvent9", m_PreEvent9, m_def_PreEvent9)
    Call PropBag.WriteProperty("PreEvent10", m_PreEvent10, m_def_PreEvent10)
    Call PropBag.WriteProperty("LockColors", m_LockColors, m_def_LockColors)
    Call PropBag.WriteProperty("LockTimeFormat", m_LockTimeFormat, m_def_LockTimeFormat)
    Call PropBag.WriteProperty("LockEventDisplay", m_LockEventDisplay, m_def_LockEventDisplay)
    Call PropBag.WriteProperty("LockMenu", m_LockMenu, m_def_LockMenu)
    Call PropBag.WriteProperty("MenuPassword", m_MenuPassword, m_def_MenuPassword)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get MilitaryTime() As Boolean
Attribute MilitaryTime.VB_Description = "Determines if the Clock Displays 24 hour time or12 hour time."
    MilitaryTime = m_MilitaryTime
End Property

Public Property Let MilitaryTime(ByVal New_MilitaryTime As Boolean)
    m_MilitaryTime = New_MilitaryTime
    PropertyChanged "MilitaryTime"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,2,0
Public Property Get SetTimeHours(SetTimeIndex As Long) As Variant
Attribute SetTimeHours.VB_MemberFlags = "400"
    If Len(dcTimerEventTriggers(SetTimeIndex)) = 6 Then
        SetTimeHours = Left(dcTimerEventTriggers(SetTimeIndex), 2)
        Else
        SetTimeHours = -1
    End If
End Property

Public Property Let SetTimeHours(SetTimeIndex As Long, ByVal New_SetTimeHours As Variant)
    If Ambient.UserMode = False Then Err.Raise 387
    Dim Testing
    Testing = New_SetTimeHours
    If Len("" & Testing) = 1 Then Testing = "0" & Testing
    If Val(Testing) < 0 Or Val(Testing) > 23 Then
        MsgBox "Invalid Input to change the Hours Setting for Index " & SetTimeIndex & ".", vbOKOnly + vbCritical, "Input Error"
        Else
        dcTimerEventTriggers(SetTimeIndex) = Testing & Mid(dcTimerEventTriggers(SetTimeIndex), 3)
    End If
    PropertyChanged "SetTimeHours"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,2,0
Public Property Get SetTimeMinutes(SetTimeIndex As Long) As Variant
Attribute SetTimeMinutes.VB_MemberFlags = "400"
    If Len(dcTimerEventTriggers(SetTimeIndex)) = 6 Then
        SetTimeMinutes = Mid(dcTimerEventTriggers(SetTimeIndex), 3, 2)
        Else
        SetTimeMinutes = -1
    End If
End Property

Public Property Let SetTimeMinutes(SetTimeIndex As Long, ByVal New_SetTimeMinutes As Variant)
    If Ambient.UserMode = False Then Err.Raise 387
    Dim Testing
    Testing = New_SetTimeMinutes
    If Len("" & Testing) = 1 Then Testing = "0" & Testing
    If Val(Testing) < 0 Or Val(Testing) > 59 Then
        MsgBox "Invalid Input to change the Minutes Setting for Index " & SetTimeIndex & ".", vbOKOnly + vbCritical, "Input Error"
        Else
        dcTimerEventTriggers(SetTimeIndex) = Left(dcTimerEventTriggers(SetTimeIndex), 2) & Testing & Mid(dcTimerEventTriggers(SetTimeIndex), 5)
    End If
    PropertyChanged "SetTimeMinutes"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,2,0
Public Property Get SetTimeSeconds(SetTimeIndex As Long) As Variant
Attribute SetTimeSeconds.VB_MemberFlags = "400"
    If Len(dcTimerEventTriggers(SetTimeIndex)) = 6 Then
        SetTimeSeconds = Right(dcTimerEventTriggers(SetTimeIndex), 2)
        Else
        SetTimeSeconds = -1
    End If
End Property

Public Property Let SetTimeSeconds(SetTimeIndex As Long, ByVal New_SetTimeSeconds As Variant)
    If Ambient.UserMode = False Then Err.Raise 387
    Dim Testing
    Testing = New_SetTimeSeconds
    If Len("" & Testing) = 1 Then Testing = "0" & Testing
    If Val(Testing) < 0 Or Val(Testing) > 59 Then
        MsgBox "Invalid Input to change the Seconds Setting for Index " & SetTimeIndex & ".", vbOKOnly + vbCritical, "Input Error"
        Else
        dcTimerEventTriggers(SetTimeIndex) = Left(dcTimerEventTriggers(SetTimeIndex), 4) & Testing
    End If
    PropertyChanged "SetTimeSeconds"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get PreEvent1() As String
Attribute PreEvent1.VB_Description = "PreEvents are events that you can set triggers for, in design mode.  The trigger must be set in an HH:MM:SS format.  Any blank or invalid PreEvents, will prevent the ones after it from being loaded."
    PreEvent1 = m_PreEvent1
End Property

Public Property Let PreEvent1(ByVal New_PreEvent1 As String)
    If Ambient.UserMode Then Err.Raise 382
    m_PreEvent1 = New_PreEvent1
    PropertyChanged "PreEvent1"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get PreEvent2() As String
Attribute PreEvent2.VB_Description = "PreEvents are events that you can set triggers for, in design mode.  The trigger must be set in an HH:MM:SS format.  Any blank or invalid PreEvents, will prevent the ones after it from being loaded."
    PreEvent2 = m_PreEvent2
End Property

Public Property Let PreEvent2(ByVal New_PreEvent2 As String)
    If Ambient.UserMode Then Err.Raise 382
    m_PreEvent2 = New_PreEvent2
    PropertyChanged "PreEvent2"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get PreEvent3() As String
Attribute PreEvent3.VB_Description = "PreEvents are events that you can set triggers for, in design mode.  The trigger must be set in an HH:MM:SS format.  Any blank or invalid PreEvents, will prevent the ones after it from being loaded."
    PreEvent3 = m_PreEvent3
End Property

Public Property Let PreEvent3(ByVal New_PreEvent3 As String)
    If Ambient.UserMode Then Err.Raise 382
    m_PreEvent3 = New_PreEvent3
    PropertyChanged "PreEvent3"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get PreEvent4() As String
Attribute PreEvent4.VB_Description = "PreEvents are events that you can set triggers for, in design mode.  The trigger must be set in an HH:MM:SS format.  Any blank or invalid PreEvents, will prevent the ones after it from being loaded."
    PreEvent4 = m_PreEvent4
End Property

Public Property Let PreEvent4(ByVal New_PreEvent4 As String)
    If Ambient.UserMode Then Err.Raise 382
    m_PreEvent4 = New_PreEvent4
    PropertyChanged "PreEvent4"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get PreEvent5() As String
Attribute PreEvent5.VB_Description = "PreEvents are events that you can set triggers for, in design mode.  The trigger must be set in an HH:MM:SS format.  Any blank or invalid PreEvents, will prevent the ones after it from being loaded."
    PreEvent5 = m_PreEvent5
End Property

Public Property Let PreEvent5(ByVal New_PreEvent5 As String)
    If Ambient.UserMode Then Err.Raise 382
    m_PreEvent5 = New_PreEvent5
    PropertyChanged "PreEvent5"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get PreEvent6() As String
Attribute PreEvent6.VB_Description = "PreEvents are events that you can set triggers for, in design mode.  The trigger must be set in an HH:MM:SS format.  Any blank or invalid PreEvents, will prevent the ones after it from being loaded."
    PreEvent6 = m_PreEvent6
End Property

Public Property Let PreEvent6(ByVal New_PreEvent6 As String)
    If Ambient.UserMode Then Err.Raise 382
    m_PreEvent6 = New_PreEvent6
    PropertyChanged "PreEvent6"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get PreEvent7() As String
Attribute PreEvent7.VB_Description = "PreEvents are events that you can set triggers for, in design mode.  The trigger must be set in an HH:MM:SS format.  Any blank or invalid PreEvents, will prevent the ones after it from being loaded."
    PreEvent7 = m_PreEvent7
End Property

Public Property Let PreEvent7(ByVal New_PreEvent7 As String)
    If Ambient.UserMode Then Err.Raise 382
    m_PreEvent7 = New_PreEvent7
    PropertyChanged "PreEvent7"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get PreEvent8() As String
Attribute PreEvent8.VB_Description = "PreEvents are events that you can set triggers for, in design mode.  The trigger must be set in an HH:MM:SS format.  Any blank or invalid PreEvents, will prevent the ones after it from being loaded."
    PreEvent8 = m_PreEvent8
End Property

Public Property Let PreEvent8(ByVal New_PreEvent8 As String)
    If Ambient.UserMode Then Err.Raise 382
    m_PreEvent8 = New_PreEvent8
    PropertyChanged "PreEvent8"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get PreEvent9() As String
Attribute PreEvent9.VB_Description = "PreEvents are events that you can set triggers for, in design mode.  The trigger must be set in an HH:MM:SS format.  Any blank or invalid PreEvents, will prevent the ones after it from being loaded."
    PreEvent9 = m_PreEvent9
End Property

Public Property Let PreEvent9(ByVal New_PreEvent9 As String)
    If Ambient.UserMode Then Err.Raise 382
    m_PreEvent9 = New_PreEvent9
    PropertyChanged "PreEvent9"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,0,
Public Property Get PreEvent10() As String
Attribute PreEvent10.VB_Description = "PreEvents are events that you can set triggers for, in design mode.  The trigger must be set in an HH:MM:SS format.  Any blank or invalid PreEvents, will prevent the ones after it from being loaded."
    PreEvent10 = m_PreEvent10
End Property

Public Property Let PreEvent10(ByVal New_PreEvent10 As String)
    If Ambient.UserMode Then Err.Raise 382
    m_PreEvent10 = New_PreEvent10
    PropertyChanged "PreEvent10"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function fSetNumbers(NumToSet As Long, GroupToSet As Long) As Variant
Dim PicNum
If NumToSet > 0 Then PicNum = NumToSet
If NumToSet = 0 Then PicNum = 10
If NumToSet = -1 Then PicNum = 11
imgNumber(GroupToSet).Picture = ilnumbers(dcColorScheme).ListImages.Item(PicNum).Picture
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get LockColors() As Boolean
Attribute LockColors.VB_Description = "Setting this to True, prevents the end users from changing the clock colors by left clicking the buttons.\r\n"
    LockColors = m_LockColors
End Property

Public Property Let LockColors(ByVal New_LockColors As Boolean)
    m_LockColors = New_LockColors
    PropertyChanged "LockColors"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get LockTimeFormat() As Boolean
Attribute LockTimeFormat.VB_Description = "Setting this to true, prevents the clock from changing from Military Time to 12 hour time when the clock numbers are right clicked.\r\n"
    LockTimeFormat = m_LockTimeFormat
End Property

Public Property Let LockTimeFormat(ByVal New_LockTimeFormat As Boolean)
    m_LockTimeFormat = New_LockTimeFormat
    PropertyChanged "LockTimeFormat"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get LockEventDisplay() As Boolean
Attribute LockEventDisplay.VB_Description = "Setting this to true, prevent the Trigger Events Display from showing when the colons are left clicked."
    LockEventDisplay = m_LockEventDisplay
End Property

Public Property Let LockEventDisplay(ByVal New_LockEventDisplay As Boolean)
    m_LockEventDisplay = New_LockEventDisplay
    PropertyChanged "LockEventDisplay"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get LockMenu() As Boolean
Attribute LockMenu.VB_Description = "Setting this to true, prevents the Main Menu from being shown, when the user right clicks the colons."
    LockMenu = m_LockMenu
End Property

Public Property Let LockMenu(ByVal New_LockMenu As Boolean)
    m_LockMenu = New_LockMenu
    PropertyChanged "LockMenu"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get MenuPassword() As String
Attribute MenuPassword.VB_Description = "This will password protect the Main Menu screen, so only you can use it to change various features during runtime."
    MenuPassword = m_MenuPassword
End Property

Public Property Let MenuPassword(ByVal New_MenuPassword As String)
    m_MenuPassword = New_MenuPassword
    PropertyChanged "MenuPassword"
End Property

