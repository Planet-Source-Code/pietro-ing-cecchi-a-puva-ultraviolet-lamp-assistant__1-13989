VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "VTEXT.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ultraviolet (PUVA) Tanning Lamp Assistant"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimedButtonConfirmTimer 
      Left            =   4560
      Top             =   4560
   End
   Begin VB.Timer DemoTimer 
      Left            =   8880
      Top             =   2760
   End
   Begin HTTSLibCtl.TextToSpeech Speaker 
      Height          =   495
      Left            =   600
      OleObjectBlob   =   "Form1.frx":058A
      TabIndex        =   6
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer SpeechTextTimer 
      Interval        =   1000
      Left            =   120
      Top             =   4440
   End
   Begin VB.Timer PhaseTimer 
      Index           =   3
      Left            =   4080
      Top             =   720
   End
   Begin VB.Timer PhaseTimer 
      Index           =   2
      Left            =   2280
      Top             =   720
   End
   Begin VB.Timer PhaseTimer 
      Index           =   1
      Left            =   480
      Top             =   720
   End
   Begin VB.Timer PhaseTimer 
      Index           =   4
      Left            =   5880
      Top             =   720
   End
   Begin VB.Timer SequenceTimer 
      Left            =   2880
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Caption         =   "configure exposure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3975
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   9015
      Begin VB.CheckBox Check1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1815
         Index           =   4
         Left            =   5640
         Picture         =   "Form1.frx":05AE
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1815
         Index           =   3
         Left            =   3840
         Picture         =   "Form1.frx":09F0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1815
         Index           =   2
         Left            =   2040
         Picture         =   "Form1.frx":0E32
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1815
         Index           =   1
         Left            =   240
         Picture         =   "Form1.frx":1274
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.PictureBox Duration 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   1
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   1515
         TabIndex        =   11
         Top             =   2400
         Width           =   1575
      End
      Begin VB.PictureBox Duration 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   4
         Left            =   5640
         ScaleHeight     =   435
         ScaleWidth      =   1515
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
      End
      Begin VB.PictureBox Duration 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   2
         Left            =   2040
         ScaleHeight     =   435
         ScaleWidth      =   1515
         TabIndex        =   9
         Top             =   2400
         Width           =   1575
      End
      Begin VB.PictureBox Duration 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   3
         Left            =   3840
         ScaleHeight     =   435
         ScaleWidth      =   1515
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Timer ShowSpeechTextTimer 
         Interval        =   500
         Left            =   720
         Top             =   3480
      End
      Begin MSForms.Label ShowSpeechText 
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   3480
         Visible         =   0   'False
         Width           =   8415
         ForeColor       =   65535
         PicturePosition =   327683
         Size            =   "14843;661"
         SpecialEffect   =   3
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.CommandButton Command1 
         Height          =   495
         Index           =   7
         Left            =   7680
         TabIndex        =   16
         ToolTipText     =   "8 minutes lasting demo"
         Top             =   2520
         Width           =   1005
         ForeColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "demo"
         Size            =   "1773;873"
         TakeFocusOnClick=   0   'False
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton Command1 
         Height          =   495
         Index           =   6
         Left            =   7680
         TabIndex        =   17
         ToolTipText     =   "loads stored configuration, if any"
         Top             =   1560
         Width           =   1005
         ForeColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "load"
         Size            =   "1773;873"
         TakeFocusOnClick=   0   'False
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton Command1 
         Height          =   495
         Index           =   5
         Left            =   7680
         TabIndex        =   18
         ToolTipText     =   "saves actual configuration"
         Top             =   720
         Width           =   1005
         ForeColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "save"
         Size            =   "1773;873"
         TakeFocusOnClick=   0   'False
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   3000
         Width           =   1575
         Size            =   "2778;661"
         Max             =   30
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   21
         Top             =   3000
         Width           =   1575
         Size            =   "2778;661"
         Max             =   30
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   20
         Top             =   3000
         Width           =   1575
         Size            =   "2778;661"
         Max             =   30
      End
      Begin MSForms.SpinButton SpinButton1 
         Height          =   375
         Index           =   3
         Left            =   3840
         TabIndex        =   19
         Top             =   3000
         Width           =   1575
         Size            =   "2778;661"
         Max             =   30
      End
   End
   Begin VB.Label UnloadMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "still exposing"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   8520
      TabIndex        =   5
      Top             =   0
      Width           =   885
   End
   Begin MSForms.CommandButton Command1 
      Height          =   630
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   6000
      Width           =   8985
      ForeColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "AKNOWLEDGE"
      Size            =   "15849;1111"
      TakeFocusOnClick=   0   'False
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton Command1 
      Height          =   435
      Index           =   4
      Left            =   1200
      TabIndex        =   3
      Top             =   6360
      Width           =   750
      VariousPropertyBits=   276824091
      Caption         =   "dummy"
      Size            =   "1323;767"
      TakeFocusOnClick=   0   'False
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Command1 
      Height          =   1455
      Index           =   2
      Left            =   6240
      TabIndex        =   2
      Top             =   4440
      Width           =   3045
      ForeColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "pause exposure"
      Size            =   "5371;2566"
      TakeFocusOnClick=   0   'False
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton Command1 
      Height          =   1455
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   4440
      Width           =   2085
      ForeColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "abort exposure"
      Size            =   "3678;2566"
      TakeFocusOnClick=   0   'False
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton Command1 
      Height          =   1455
      Index           =   0
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "wear protective goggles!"
      Top             =   4440
      Width           =   2445
      ForeColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "start exposure"
      Size            =   "4313;2566"
      Picture         =   "Form1.frx":16B6
      TakeFocusOnClick=   0   'False
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Image BackGroundImage 
      Height          =   375
      Index           =   1
      Left            =   480
      Top             =   6360
      Width           =   375
      BorderStyle     =   0
      Size            =   "661;661"
      Picture         =   "Form1.frx":1F90
      PictureTiling   =   -1  'True
      VariousPropertyBits=   19
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Targetphasecount(1 To 4) As Integer 'seconds
Dim PhaseNotBusy As Boolean
Dim RunningPhase As Integer
Dim SequenceRunning As Boolean
Dim AbortSequence As Boolean
Dim AbortPhase As Boolean
Dim Clickedalready As Boolean

Enum DisplayType
     DisplayPercent = 1
     DisplayMessage = 2
End Enum

Dim SpeechText As String
Dim SpeechTextHowManyTimes As Integer
'NOTE:
'Speaker.Speaking= 0(not speaking) or 1(speaking)
'Speaker.IsSpeaking= 0(not speaking) or 1(speaking)
Dim Xtxt() As String

Private Sub Check1_Click(Index As Integer)
 Select Case Check1(Index).Value
   Case 0 'unchecked
      For a = 1 To 4
       If Check1(a).Caption <> "" Then
          If Check1(a).Caption > Val(Check1(Index).Caption) Then
            Check1(a).Caption = Val(Check1(a).Caption) - 1
          End If
       End If
      Next
      Check1(Index).Caption = ""
      Check1(Index).BackColor = vbButtonFace
   Case 1 'checked
      Max = 0
      For a = 1 To 4
       If Val(Check1(a).Caption) > Max Then
         Max = Val(Check1(a).Caption)
       End If
      Next
      Check1(Index).Caption = Max + 1
      Check1(Index).BackColor = vbCyan
   Case 2 'grayed
     'not used
     Stop
 End Select
 
 'this to avoid disturbing tab selection
 'indication on control
' Command1(IIf(Command1(0).Enabled, 0, 1)).SetFocus
If Enabled Then Command1(4).SetFocus

 End Sub

Private Sub Command1_Click(Index As Integer)
 Select Case Index
  Case 0 'start
     'commands
     Command1(0).Enabled = False 'start
     Command1(1).Enabled = True  'abort
     Command1(2).Enabled = False 'pause, will be enabled in SequenceTimer_Timer
     Command1(4).SetFocus 'dummy command, just to get focus
     'indicators
     For a = 1 To 4
        pct = 0
        msg = CStr(SpinButton1(a).Value) & " min"
        UpdateProgress Duration(a), pct, msg, DisplayMessage
     Next
     'start sequence
     For a = 1 To 4
       PhaseTimer(a).Enabled = False
       PhaseTimer(a).Interval = 1000
     Next
     'inhibit configure
     Frame1.Enabled = False
     SequenceTimer.Interval = 50
     SequenceTimer.Enabled = True
     SequenceTimer_Timer 'start immediately
  Case 1 'abort
     'Clickedalready is a boolean global used for aborting sequence
     If Not Clickedalready Then
        Clickedalready = True
        TimedButtonConfirmTimer_Timer 'uses Clickedalready
     Else
        'commands
        Command1(0).Enabled = True 'start
        Command1(1).Enabled = False  'abort
        Command1(2).Enabled = False  'pause
        Command1(3).Enabled = False  'pause
        Command1(4).SetFocus
        'globals
        AbortPhase = True
        AbortSequence = True
        'reset
        Command1(1).Caption = "abort exposure"
        Clickedalready = False
   
     End If
  Case 2 'pause
     PhaseTimer(RunningPhase).Enabled = Not PhaseTimer(RunningPhase).Enabled
     Command1(1).Enabled = PhaseTimer(RunningPhase).Enabled
     Command1(2).BackColor = IIf(PhaseTimer(RunningPhase).Enabled, vbButtonFace, &HC0FFFF) 'vbYellow
     Command1(2).Caption = IIf(PhaseTimer(RunningPhase).Enabled, "pause exposure", "EXPOSURE PAUSED," & vbNewLine & "PHASE " & Check1(RunningPhase).Caption)
     Command1(3).Enabled = False  'pause
  Case 3 'aknowledge
     'commands
     Command1(0).Enabled = True 'start
     Command1(1).Enabled = False  'abort
     Command1(2).Enabled = False  'pause
     Command1(3).BackColor = vbButtonFace 'aknowledge  'was &H0080C0FF& apricot
     Command1(3).Enabled = False  'pause
     Command1(3).Caption = "ACKNOWLEDGE"
     Command1(4).SetFocus
     SpeechTextStop
  Case 5 'save configuration to registry
 
     Frame1.Enabled = False
     Frame1.BackColor = vbYellow
     
     On Error Resume Next
        DeleteSetting App.EXEName
     On Error GoTo 0
   
     'save check1(1-4) and spinbutton1(1-4)
     For a = 1 To 4
        SaveSetting App.EXEName, "CHECK1", "Check1(" & a & ").Caption", CStr(Check1(a).Caption)
        SaveSetting App.EXEName, "SPINBUTTON1", "Spinnbutton1(" & a & ").Value", CStr(SpinButton1(a).Value)
        Pause 300
     Next
     
     Frame1.BackColor = vbButtonFace
     Frame1.Enabled = True
  
  Case 6 'load configuration from registry
     
     Frame1.Enabled = False
     Frame1.BackColor = vbYellow
    
     For a = 1 To 4
        SpinButton1(a).Value = 1
        SpinButton1_Change (a)
        Pause 100
     Next
     
     For a = 1 To 4
        ret = GetSetting(App.EXEName, "SPINBUTTON1", "Spinnbutton1(" & a & ").Value", "no value")
        If ret <> "no value" Then
           SpinButton1(a).Value = ret
           SpinButton1_Change (a)
           Pause 100
        End If
     Next
  
     For a = 1 To 4
        Check1(a).Value = vbUnchecked
        Check1_Click Val(a)
        Pause 100
     Next
     
     incr = 1
     For aa = 1 To 4
     For a = 1 To 4
        capt = GetSetting(App.EXEName, "CHECK1", "Check1(" & a & ").Caption", "no value")
        If capt <> "no value" Then
            If Val(capt) = incr Then
                Check1(a).Value = vbChecked
                Check1(a).Caption = capt
          '      Check1_Click Val(capt)
                Pause 100
                incr = incr + 1
            End If
        End If
     Next
     Next
  
     Frame1.BackColor = vbButtonFace
     Frame1.Enabled = True
  
  Case 7 'demo all operations
 
     DemoTimer_Timer
     
 
 End Select
 
End Sub

Private Sub DemoTimer_Timer()
 'runs the demo
 Static pass As Integer
 
   DemoTimer.Enabled = False
   DemoTimer.Interval = 100
 If Speaker.IsSpeaking = 1 Then
   DemoTimer.Interval = 500
   DemoTimer.Enabled = True
   Exit Sub
 End If
 pass = pass + 1
 Select Case pass
  Case 1 'assign txt
     
  'debug
     Enabled = False
     
     Command1(7).BackColor = vbGreen
     ReDim Xtxt(100)
     Xtxt(2) = "Welcome."
     Xtxt(3) = "This is the little demo of Ultraviolet Lamp Assistant."
     Xtxt(4) = "I will force some commands to let you see"
     Xtxt(5) = "the program working."
     Xtxt(6) = "First of all I will reset all configuration controls"
     
     Xtxt(8) = "O.K."
     Xtxt(9) = "Now I start clicking on the front side exposure"
     Xtxt(10) = "Yes, the full moon!"
     
     Xtxt(12) = "Number 1 in the button means this is phase 1 of exposure"
     Xtxt(13) = "Now I increase the phase time to 15 minutes"
     
     Xtxt(15) = "Now I do the same for rear side, id est empty moon phase"
     
     Xtxt(17) = "And for the left side"
     
     Xtxt(19) = "And for the right side"
     
     Xtxt(21) = "Note the order we clicked on the phases"
     Xtxt(22) = "Now suppose I changed mind on full moon phase"
     Xtxt(23) = "Suppose I will do the front exposure last"
     Xtxt(24) = "With a click I reset full moon"
     
     Xtxt(26) = "Note that all other phases reorder automatically"
     Xtxt(27) = "Now a second click wil set full moon as the fourth phase"
          
     Xtxt(29) = "Now let's start the exposure"
     Xtxt(30) = "NOTE: you are supposed, out of the demo, to wear"
     Xtxt(31) = "the protective goggles and to turn on the lamp"
     Xtxt(32) = "Don't do this now, of course"
     Xtxt(33) = "So, now I start the exposure sequence"
  
     Xtxt(35) = "Wait, patiently, for the indicator of phase 1"
     Xtxt(36) = "to show an increment"
     
     Xtxt(38) = "The sequence is goin on correctly."
     Xtxt(39) = "The phase will take 15 minutes to complete."
     Xtxt(40) = "We need to pause the exposure? let's do it."
     
     Xtxt(42) = "And now, let's resume from pause."
     
     Xtxt(44) = "Takes a long time, isn't?"
     Xtxt(45) = "Hence, let's abort the simulated exposure."
         
     Xtxt(47) = "Note that the button is timed, while asking you to confirm."
     Xtxt(48) = "Once expired, the abort command has been ignored."
     Xtxt(49) = "Now, retry the abort command, confirming it, instead."
     
     Xtxt(51) = "This time the command has been accepted"
     Xtxt(52) = "and the control panel now asks for aknowledge"
     Xtxt(53) = "Before aknowledging the command, we can"
     Xtxt(54) = "take a last glance to the phases indicators"
     Xtxt(55) = "It came now the time to aknowledge the closed sequence"
  
     Xtxt(57) = "We could now restart the exposure"
     Xtxt(58) = "or, better, shorten the phases time to 1 minute"
     
     Xtxt(60) = "Oh! well! let's restart now the exposure"
     Xtxt(61) = "Are you sitting comfortably and relaxing?"
     Xtxt(62) = "The show begins... and I am muting myself"
     Xtxt(63) = "From now on, up to exposure completion"
     Xtxt(64) = "audio messages are programmatic"
     Xtxt(65) = "I'll see you later"
     
     Xtxt(67) = "O.K., here I am, back to you"
     Xtxt(68) = "Did you like the show?"
     Xtxt(69) = "And the demo?"
     Xtxt(70) = "You can change it at any time, because you have the code!"
     Xtxt(71) = "A last thing to finish: please vote sincerely this application"
     Xtxt(72) = "A big hug and God bless you."
     more = ""
     Select Case Month(Date)
       Case 3, 4
         more = "Good Easter"
       Case 12
         more = "Merry Christmas"
       Case 1
         more = "Happy new Year"
     End Select
     Xtxt(72) = Xtxt(72) & " " & more & " " & Year(Date)
     
     Xtxt(73) = "goodbye"
  
  Case 2 To 6, 8 To 10, 12, 13, 15, 17, 19, 21 To 24, 26, 27, 29 To 33, 35, 36, 38 To 40, 42, 44, 45, 47 To 49, 51 To 55, 57, 58, 60 To 65, 67 To 73
     SpeechTextHowManyTimes = 0
     SpeechText = Xtxt(pass)
     Speaker.Speak Xtxt(pass)
  Case 7 'reset commands
     Frame1.BackColor = vbYellow
     For a = 1 To 4
        SpinButton1(a).Value = 1
        Pause 100
     Next
     For a = 1 To 4
        Check1(a).Value = vbUnchecked
        Pause 100
     Next
     Frame1.BackColor = vbButtonFace
  Case 11 'clik the full moon
     Check1(4).Value = vbChecked
  Case 14 'full moon to 15 minutes
     SpinButton1(4).Value = 15
  Case 16 'empty moon set
     Check1(1).Value = vbChecked
     SpinButton1(1).Value = 15
  Case 18 'left side
     Check1(2).Value = vbChecked
     SpinButton1(2).Value = 15
  Case 20 'right side
     Check1(3).Value = vbChecked
     SpinButton1(3).Value = 15
  Case 25 'deselect full moon from 4th place
     Check1(4).Value = vbUnchecked
  Case 28 'reselect full moon to 1st place
     Check1(4).Value = vbChecked
  Case 34 'start exposure
     Enabled = True
        Command1_Click 0
     Enabled = False
  Case 37 'wait indicator moves
     Pause 2000
  Case 41, 43 'pausing exposure, resuming from pause
     Enabled = True
        Command1_Click 2
     Enabled = False
  Case 46 'abort exposure
     Enabled = True
        Command1_Click 1
     Enabled = False
  Case 50 'abort and confirm
     Enabled = True
        Command1_Click 1
     Enabled = False
     Pause 1000
     Enabled = True
        Command1_Click 1
     Enabled = False
  Case 56 'aknowledge
     Enabled = True
        Command1_Click 3
     Enabled = False
  Case 59 'shorten phases duration to 1 minute
     For a = 1 To 4
       SpinButton1(a).Value = 1
       DoEvents
     Next
  Case 66 'looking at the show
     Enabled = True
     If Command1(0).Enabled Then
           Command1_Click 0
     End If
     Enabled = False
     Pause 500
     Enabled = True
     If Not Command1(3).Enabled Then
        pass = pass - 1 'wait until and of exposure aknowledge
     Else 'aknowledge end of sequence
        Enabled = False
        Pause 5000
        Enabled = True
           Command1_Click 3
        Enabled = False
        Pause 500
     End If
     Enabled = False
  
  Case Else
     SpeechTextHowManyTimes = 0
     SpeechText = ""
     DemoTimer.Interval = 0
     DemoTimer.Enabled = False
     Command1(7).BackColor = vbButtonFace
     Enabled = True
     Exit Sub
 End Select
 DemoTimer.Interval = 100
 DemoTimer.Enabled = True
   
End Sub



'Private Sub CommandButton1_DblClick(Cancel As MSForms.ReturnBoolean)
'  Cancel = True
'  CommandButton1.Enabled = False
'End Sub

Private Sub Command1_DblClick(Index As Integer, Cancel As MSForms.ReturnBoolean)
  'avoid control lock down on double click
  Cancel = True
End Sub

Private Sub Form_Load()
  Command1(0).Enabled = True
  Command1(1).Enabled = False
  Command1(2).Enabled = False
  Command1(3).Enabled = False

  For a = 1 To 4
   SpinButton1(a).Value = 1
   pct = 0
   msg = CStr(SpinButton1(a).Value) & " min"
   UpdateProgress Duration(a), pct, msg, DisplayMessage
  Next
  PhaseNotBusy = True
  Show
  
  Command1(4).Left = -2 * Command1(4).Width
  Command1(4).SetFocus 'dummy, to get focus

  BackGroundImage(1).PictureTiling = True
  BackGroundImage(1).Move 0, 0, ScaleWidth, ScaleHeight
  BackGroundImage(1).Enabled = False
  
 
  UnloadMessage.Visible = False

  Width = 9600
  'access taskbar:
  Height = 7200 - 2 * Screen.TwipsPerPixelY

' SpeechTextHowManyTimes = 1
' SpeechText = "Welcome"
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Cancel = SequenceRunning
' If Cancel Then MsgBox "Can't exit: sequence is running"
 If Cancel Then
  UnloadMessage.Visible = True
   Pause 2000
   UnloadMessage.Visible = False
 End If
End Sub

Private Sub PhaseTimer_Timer(Index As Integer)
 Static phasecount(1 To 4) As Integer
 
 If AbortPhase Then
     phasecount(Index) = 0
     PhaseTimer(Index).Enabled = False
     PhaseNotBusy = True
     If SpinButton1(Index).Value = 0 Then
        If phasecount(Index) = 0 Then pct = 0 Else pct = 100
     Else
        pct = phasecount(Index) / (60 * SpinButton1(Index).Value) * 100
     End If
     AbortPhase = False
     Exit Sub
 End If
 
 
 Select Case phasecount(Index) 'each second
  Case 0 'some initialize, starting phase
     'this prephase doesn't take time
     'because skips to go_on
     PhaseNotBusy = False
     phasecount(Index) = phasecount(Index) + 1
     If Targetphasecount(Index) = 0 Then GoTo go_end
     pct = phasecount(Index) / (60 * SpinButton1(Index).Value) * 100
  Case Is >= Targetphasecount(Index) 'ending phase
go_end:
     phasecount(Index) = 0
     PhaseTimer(Index).Enabled = False
     PhaseNotBusy = True
     pct = 100 'phasecount(Index) / (60 * SpinButton1(Index).Value) * 100
     msg = CStr(SpinButton1(Index).Value) & " min"
     UpdateProgress Duration(Index), pct, msg, DisplayMessage
     Exit Sub
  Case Else 'phase in progress
     pct = phasecount(Index) / (60 * SpinButton1(Index).Value) * 100
     msg = CStr(SpinButton1(Index).Value) & " min"
     UpdateProgress Duration(Index), pct, msg, DisplayMessage
     phasecount(Index) = phasecount(Index) + 1
 End Select
 
  
End Sub

Private Sub SequenceTimer_Timer()
'this module executes the sequence
 Static pass As Integer
 Static totalpass As Integer
 Static indirect(1 To 4) As Integer
 
 If AbortSequence Then
     PhaseNotBusy = True
     RunningPhase = 0
     Command1(0).Enabled = False
     Command1(1).Enabled = False
     Command1(2).Enabled = False
     SequenceTimer.Enabled = False
     Command1(3).BackColor = &H80C0FF   'apricot
     Command1(3).Caption = "AKNOWLEDGE: exposure aborted by user"
     Command1(3).Enabled = True
     Do
        DoEvents
     Loop Until Command1(3).Enabled = False
     Command1(0).Enabled = True
     Command1(1).Enabled = False
     Command1(2).Enabled = False
     For a = 1 To 4
      pct = 0
      msg = CStr(SpinButton1(a).Value) & " min"
      UpdateProgress Duration(a), pct, msg, DisplayMessage
     Next
     Frame1.Enabled = True
     pass = 0
     AbortSequence = False
     SequenceRunning = False
     Exit Sub
 End If
 
 Select Case pass
   Case 0 'initialize
       SequenceRunning = True
       For a = 1 To 4
          indirect(a) = 0
       Next
       totalpass = 0
       For a = 1 To 4
        If Check1(a).Caption <> "" Then
          indirect(Val(Check1(a).Caption)) = a
          Command1(2).Enabled = True 'pause command
          totalpass = totalpass + 1
        End If
       Next

       pass = pass + 1
   Case 1 To 4
     If PhaseNotBusy Then
       If indirect(pass) > 0 Then
          If Check1(indirect(pass)).Value = vbChecked Then
             RunningPhase = indirect(pass)
             Command1(2).Enabled = True
             StartPhase indirect(pass)
             SpeechMessagePhase pass, totalpass, indirect(pass)
          End If
       Else
          RunningPhase = 0
          Command1(2).Enabled = False
       End If
       pass = pass + 1
     End If
   Case 5 'close sequence
     If PhaseNotBusy Then
        RunningPhase = 0
        Command1(1).Enabled = False
        Command1(2).Enabled = False
        Command1(3).BackColor = &H80C0FF   'apricot
        Command1(3).Caption = "AKNOWLEDGE: exposure complete"
        Command1(3).Enabled = True
        SpeechMessageEndOfPhase
             
        pass = pass + 1
     End If
   Case 6 'close sequence, continued
     'wait end of sequence aknowlwdged
     If Not Command1(3).Enabled Then
        SequenceTimer.Enabled = False
        Command1(0).Enabled = True
        Command1(1).Enabled = False
        Command1(2).Enabled = False
        For a = 1 To 4
         pct = 0
         msg = CStr(SpinButton1(a).Value) & " min"
         UpdateProgress Duration(a), pct, msg, DisplayMessage
        Next
        Frame1.Enabled = True
        pass = 0
        SequenceRunning = False
     End If
   
   Case Else
    'not applicable
    Stop
 End Select
 



End Sub

Private Sub StartPhase(ByVal pass As Integer)
 PhaseNotBusy = False
 'targetphasecount in seconds
 Targetphasecount(pass) = 60 * SpinButton1(pass).Value
 PhaseTimer(pass).Interval = 1000
 PhaseTimer(pass).Enabled = True
 PhaseTimer_Timer pass 'start immediately
End Sub

'example call
'UpdateProgress Today, TodayPercent, TodayMessage, DisplayMessage
'TodayPercentSave = TodayPercent
'where
'Enum DisplayType
'     DisplayPercent = 1
'     DisplayMessage = 2
'End Enum
'Dim TodayMessage As String
'Dim TodayPercent As Integer
'
'Enum TodayTimerTextType
'     NormalText = 1
'     NoExpirePauseText = 2
'     ExitText = 3
'End Enum
'Dim TodayTimerText As TodayTimerTextType
Sub UpdateProgress(pb As Control, ByVal percent, ByVal message, ByVal disptype As DisplayType)

    Dim num As String 'use percent
    Dim msg As String 'use massage

    If Not pb.AutoRedraw Then 'picture in memory ?
        pb.AutoRedraw = -1 'no, make one
    End If

    pb.Cls 'clear picture in memory
    pb.ScaleWidth = 100 'new scalemodus
    pb.DrawMode = 10 'not XOR Pen Modus

    If IsNull(percent) Or (percent = "") Then percent = 0
    
    num = Format(percent, "###") + "%"
    
    
    msg = Choose(disptype, num, message)
    
    pb.CurrentX = 50 - pb.TextWidth(msg) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(msg)) / 2
    pb.Print msg 'print percent
    pb.Line (0, 0)-(percent, pb.ScaleHeight), , BF
    pb.Refresh 'show differents
End Sub

Private Sub ShowSpeechTextTimer_Timer()
  If SpeechText <> "" And (Speaker.IsSpeaking = 1) Then
     ShowSpeechText.Visible = True
     ShowSpeechText.Caption = "AUDIO:  " & SpeechText
  Else
     ShowSpeechText.Visible = False
  End If
End Sub

Private Sub SpeechTextTimer_Timer()
 Static times As Integer
 
 SpeechTextTimer.Enabled = False
 SpeechTextTimer.Interval = 100
 
 If SpeechText <> "" Then
    times = times + 1
    If times <= SpeechTextHowManyTimes Then
      
      'speak
      On Error Resume Next
        'DirectSS
        speaks = Speaker.Speaking
        'TextToSpeech
        speaks = Speaker.IsSpeaking
      On Error GoTo 0
      If speaks = 0 Then 'not speaking=0, speaking=1
         'asyncro request
         Speaker.Speak SpeechText
      Else
         'no increment
         times = times - 1
      End If
      SpeechTextTimer.Enabled = True
    Else 'end of repeated message,reset variables
      SpeechText = ""
      SpeechTextHowManyTimes = 0
      times = 0
    End If
 End If
 
End Sub

Private Sub SpinButton1_Change(Index As Integer)
   pct = 0
   msg = CStr(SpinButton1(Index).Value) & " min"
   UpdateProgress Duration(Index), pct, msg, DisplayMessage
   If SpinButton1(Index).Value = 0 Then Check1(Index).Value = vbUnchecked
End Sub

Private Sub Pause(ByVal milliseconds As Double) 'milliseconds
Dim Start, Finish, TotalTime
    Start = Timer * 1000 ' Set start time.
    Do While Timer * 1000 < Start + milliseconds
        DoEvents    ' Yield to other processes.
        If Timer * 1000 < Start Then 'trepassing midnight
          Start = Start - 24# * 60 * 60 * 1000
        End If
    Loop
End Sub

Private Sub SpeechMessagePhase(ByVal pass, ByVal totalpass, ByVal indir)
  'phase with no time, skip message
  If SpinButton1(indir).Value = 0 Then Exit Sub
  
  'prepare text
  side = Choose(indir, "rear", "left", "right", "front")
  txt = side & ", " & side & ", "
  txt = txt & "phase " & pass & ", of " & totalpass
  SpeechText = txt
  SpeechTextHowManyTimes = 3
  'launch speaking
  SpeechTextTimer.Enabled = True
End Sub

Private Sub SpeechMessageEndOfPhase()
  'prepare text
  txt = "exposure complete, complete. shut off ultraviolet lamp"
  SpeechText = txt
  SpeechTextHowManyTimes = 1000 'continues alerting, stopped in Command1(3) Aknowledge
  'launch speaking
  SpeechTextTimer.Enabled = True
End Sub

Private Sub SpeechTextStop()
'  Speaker.AudioPause
'  SpeechTextTimer.Enabled = False
  SpeechTextHowManyTimes = 0
'  'launch speaking
'  Speaker.AudioResume
End Sub

Private Sub TimedButtonConfirmTimer_Timer()

 Static expiring As Integer
 Static toggle As Boolean
  
  expiring = expiring - 1
  toggle = Not toggle
  If expiring < 0 Then expiring = 20: toggle = True 'start countdown
  
  TimedButtonConfirmTimer.Enabled = False
  TimedButtonConfirmTimer.Interval = 300
  
  
  If (Not Clickedalready) Or (expiring = 0) Then
     Command1(1).Caption = "abort exposure"
     Clickedalready = False
     TimedButtonConfirmTimer.Enabled = False
     TimedButtonConfirmTimer.Interval = 0
     Exit Sub
  End If
  Command1(1).Caption = IIf(toggle, "confirm " & expiring / 2, "")
  TimedButtonConfirmTimer.Enabled = True

End Sub
