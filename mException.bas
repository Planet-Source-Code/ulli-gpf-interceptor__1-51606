Attribute VB_Name = "mException"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const EXCEPTION_MAXIMUM_PARAMETERS  As Long = 15
Private Const EXCEPTION_EXECUTE_HANDLER     As Long = 1
Private Const EXCEPTION_CONTINUE_EXECUTION  As Long = -1

Private Type EXCEPTION_POINTERS
    pExceptionRecord    As Long 'pointer to an EXCEPTION_RECORD structure
    pContextRecord      As Long 'pointer to a CONTEXT structure
End Type

'not used, just for documentation
'Private Type EXCEPTION_RECORD
'    ExceptionCode       As Long ' + 0
'    ExceptionFlags      As Long ' + 4
'    pExceptionRecord    As Long ' + 8    pointer to a nested EXCEPTION_RECORD structure
'    ExceptionAddress    As Long ' + 12
'    NumberParameters    As Long ' + 16
'    ExceptionInformation(0 To EXCEPTION_MAXIMUM_PARAMETERS) As Long
'End Type

Public myNoisy          As Boolean

Public Function Interceptor(lpException As EXCEPTION_POINTERS) As Long

  Dim i As Long, j As Long

    With lpException
        CopyMemory i, ByVal .pExceptionRecord + 0, 4 'exception code
        CopyMemory j, ByVal .pExceptionRecord + 12, 4 'exception address
    End With 'LPEXCEPTION
    With fException
        .txRef = Hex$(i) & " (" & i & ")" & " / " & Format$(Right$("00000000" & Hex$(j), 8), "<@@\-@@\-@@\-@@") & " (" & Format$(j, "#,0") & ")"
        Do While .Visible
            DoEvents
        Loop
        If .Tag = "1" Then 'continue
            Interceptor = EXCEPTION_CONTINUE_EXECUTION
          Else 'NOT .TAG...
            Interceptor = EXCEPTION_EXECUTE_HANDLER
        End If
    End With 'FEXCEPTION
    Unload fException
    Set fException = Nothing

End Function

':) Ulli's VB Code Formatter V2.16.14 (2004-Feb-09 00:11) 24 + 26 = 50 Lines
