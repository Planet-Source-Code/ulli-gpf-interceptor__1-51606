VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "fTest"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton btFake 
      BackColor       =   &H008080FF&
      Caption         =   "Cause Exception"
      Height          =   615
      Index           =   1
      Left            =   2100
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "without interceptor"
      Top             =   345
      Width           =   1395
   End
   Begin VB.CommandButton btFake 
      BackColor       =   &H0000C000&
      Caption         =   "Cause Exception"
      Height          =   615
      Index           =   0
      Left            =   420
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "with interceptor"
      Top             =   345
      Width           =   1395
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'used to fake an exception in the IDE
Private Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

'our exception handler
Private ExcHdlr     As cException

Private Sub btFake_Click(Index As Integer)

  Dim i(0 To 1) As Long

    Select Case Index

      Case 0 'raise exception with active interceptor
        Set ExcHdlr = New cException    'set own exception handler active
        ExcHdlr.Noisy = True            'switch on noise
        'i(555555) = 1                   'true exception - index out of range (when compiled)
        RaiseException 55, 0, 0, 0      'or - fake an exception (for testing in the IDE)
        Set ExcHdlr = Nothing           'set own exception handler inactive

      Case Else 'raise exception with inactive interceptor
        'i(555555) = 1                   'true exception - index out of range (when co,piled)
        RaiseException 55, 0, 0, 0      'or - fake an exception (for testing in the IDE)

    End Select

End Sub

Private Sub Form_Load()

    Move (Screen.Width - Width) / 5, (Screen.Height - Height) / 5

End Sub

':) Ulli's VB Code Formatter V2.16.14 (2004-Feb-09 00:11) 7 + 29 = 36 Lines
