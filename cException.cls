VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Private PrevProcPtr As Long

Private Sub Class_Initialize()

  'set interceptor active

    PrevProcPtr = SetUnhandledExceptionFilter(Far(AddressOf Interceptor)) 'hihi - cheating VB

End Sub

Private Sub Class_Terminate()

  'set interceptor back to what it was

    SetUnhandledExceptionFilter PrevProcPtr

End Sub

Private Function Far(ByVal Pointer As Long) As Long

  'little helper to be able to use 'AddressOf' in a Class

    Far = Pointer

End Function

Public Property Let Noisy(nuNoisy As Boolean)

    myNoisy = CBool(nuNoisy)

End Property

Public Property Get Noisy() As Boolean

    Noisy = myNoisy

End Property

':) Ulli's VB Code Formatter V2.16.14 (2004-Feb-09 00:11) 4 + 38 = 42 Lines
