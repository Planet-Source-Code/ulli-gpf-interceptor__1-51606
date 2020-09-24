VERSION 5.00
Begin VB.Form fException 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'Kein
   Caption         =   "Sorry..."
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txRef 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1890
      Width           =   4470
   End
   Begin VB.CommandButton btTC 
      BackColor       =   &H008080FF&
      Caption         =   "Terminate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3360
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   2370
      Width           =   1095
   End
   Begin VB.CommandButton btTC 
      BackColor       =   &H0000C000&
      Caption         =   "Continue"
      Height          =   375
      Index           =   1
      Left            =   4575
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   2370
      Width           =   1095
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "terminate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000E0&
      Height          =   195
      Index           =   5
      Left            =   210
      TabIndex        =   8
      Top             =   1260
      Width           =   795
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000A000&
      Height          =   195
      Index           =   4
      Left            =   1485
      TabIndex        =   7
      Top             =   1065
      Width           =   750
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Reference:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00686800&
      Height          =   195
      Index           =   2
      Left            =   210
      TabIndex        =   6
      Top             =   1920
      Width           =   960
   End
   Begin VB.Shape shp 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Height          =   2895
      Left            =   30
      Top             =   30
      Width           =   5835
   End
   Begin VB.Line ln 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   990
      X2              =   4890
      Y1              =   690
      Y2              =   705
   End
   Begin VB.Line ln 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   1515
      X2              =   5415
      Y1              =   225
      Y2              =   240
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "strongly"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   3030
      TabIndex        =   5
      Top             =   1065
      Width           =   675
   End
   Begin VB.Image img 
      Appearance      =   0  '2D
      Height          =   570
      Left            =   210
      Picture         =   "fException.frx":0000
      Top             =   165
      Width           =   570
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   $"fException.frx":117A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00686800&
      Height          =   795
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   870
      Width           =   5535
   End
   Begin VB.Label lb 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Sorry, a Fatal Error has occurred"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   1245
      TabIndex        =   2
      Top             =   330
      Width           =   3915
   End
End
Attribute VB_Name = "fException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (lpszName As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Const SND_ASYNC     As Long = 1
Private Const SND_NODEFAULT As Long = 2
Private Const SND_MEMORY    As Long = 4

Private ToolTips(1 To 4)    As cToolTip

Private Sub btTC_Click(Index As Integer)

    Tag = Index  'Tag is used to pass back what the user decided
    Hide

End Sub

Private Sub Form_Load()

  Dim i         As Long
  Dim Sound()   As Byte

    Set ToolTips(1) = New cToolTip
    Set ToolTips(2) = New cToolTip
    Set ToolTips(3) = New cToolTip
    ToolTips(1).Create txRef, "        ¯¯¯¯¯¯¯¯¯¯¯¯|        Error Code / Memory Address||        Click to copy to clipboard.", TTBalloonIfActive, True, TTIconInfo, "Error Details", &HA0&, &HD8F0FF, 50, 20000
    ToolTips(2).Create btTC(0), "Terminates the Application", TTBalloonIfActive, False, TTIconInfo, "Terminate", &HA0&, &HD8F0FF
    ToolTips(3).Create btTC(1), "Attempts to continue the Application", TTBalloonIfActive, False, TTIconInfo, "Continue", &HA0&, &HD8F0FF

    Show
    'noisy and this little gimmick could be removed alltogether if you don't like it
    If myNoisy Then
        DoEvents 'let the form paint
        Sound = LoadResData("Crash", "Sound")
        If PlaySound(Sound(0), 0, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY) Then 'play sound
            Sleep 2400 'wait until crash sounds
            Rnd -6 '(-6) produces a good balanced random sequence
            For i = 360 To 0 Step -3 'shake form
                DoEvents
                Sleep 30
                Move Left + (0.5 - Rnd) * i, Top + (0.5 - Rnd) * i
            Next i
        End If
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  'the mouse is outside the windowless control

    If Not ToolTips(4) Is Nothing Then                 'and there is a tooltip
        Set ToolTips(4) = Nothing                      'and destroy the class instance
    End If

End Sub

Private Sub lb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If Index = 1 Then
        If ToolTips(4) Is Nothing Then                     'we have no tooltip yet for this
            Set ToolTips(4) = New cToolTip                 'instantiate the class
            ToolTips(4).Create lb(1), "A fatal error was detected by the Windows|environment and the application has been|put in a halted state, pending abortion.", , , TTIconError, "Fatal Error", &HA0&, &HD8F0FF, , 15000 'create a tooltip for the containing form instead
        End If
    End If

End Sub

Private Sub txRef_Click()

    Clipboard.Clear
    Clipboard.SetText txRef

End Sub

Private Sub txRef_GotFocus()

    btTC(0).SetFocus

End Sub

Private Sub Shakeform()

End Sub

':) Ulli's VB Code Formatter V2.16.14 (2004-Feb-09 00:11) 9 + 77 = 86 Lines
