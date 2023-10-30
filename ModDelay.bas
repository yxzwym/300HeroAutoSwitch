Attribute VB_Name = "ModDelay"
' ÑÓ³ÙÄ£¿é

' ¹Ì¶¨ÑÓ³Ù
' Delay(ms as long)

' Ëæ»úÑÓ³Ù
' Delay(min as long, max as long)

Option Explicit

Private Type FILETIME

dwLowDateTime As Long
dwHighDateTime As Long

End Type

Private Const WAIT_ABANDONED& = &H80&

Private Const WAIT_ABANDONED_0& = &H80&

Private Const WAIT_FAILED& = -1&

Private Const WAIT_IO_COMPLETION& = &HC0&

Private Const WAIT_OBJECT_0& = 0

Private Const WAIT_OBJECT_1& = 1

Private Const WAIT_TIMEOUT& = &H102&

Private Const INFINITE = &HFFFF

Private Const ERROR_ALREADY_EXISTS = 183&

Private Const QS_HOTKEY& = &H80

Private Const QS_KEY& = &H1

Private Const QS_MOUSEBUTTON& = &H4

Private Const QS_MOUSEMOVE& = &H2

Private Const QS_PAINT& = &H20

Private Const QS_POSTMESSAGE& = &H8

Private Const QS_SENDMESSAGE& = &H40

Private Const QS_TIMER& = &H10

Private Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)

Private Const QS_INPUT& = (QS_MOUSE Or QS_KEY)

Private Const QS_ALLEVENTS& = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)

Private Const QS_ALLINPUT& = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Private Const UNITS = 4294967296#

Private Const MAX_LONG = -2147483648#

Private Declare Function CreateWaitableTimer _
Lib "kernel32" _
Alias "CreateWaitableTimerA" (ByVal lpSemaphoreAttributes As Long, _
ByVal bManualReset As Long, _
ByVal lpName As String) As Long

Private Declare Function OpenWaitableTimer _
Lib "kernel32" _
Alias "OpenWaitableTimerA" (ByVal dwDesiredAccess As Long, _
ByVal bInheritHandle As Long, _
ByVal lpName As String) As Long

Private Declare Function SetWaitableTimer _
Lib "kernel32" (ByVal hTimer As Long, _
lpDueTime As FILETIME, _
ByVal lPeriod As Long, _
ByVal pfnCompletionRoutine As Long, _
ByVal lpArgToCompletionRoutine As Long, _
ByVal fResume As Long) As Long

Private Declare Function CancelWaitableTimer Lib "kernel32" (ByVal hTimer As Long)

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function WaitForSingleObject _
Lib "kernel32" (ByVal hHandle As Long, _
ByVal dwms As Long) As Long

Private Declare Function MsgWaitForMultipleObjects _
Lib "user32" (ByVal nCount As Long, _
pHandles As Long, _
ByVal fWaitAll As Long, _
ByVal dwms As Long, _
ByVal dwWakeMask As Long) As Long

Private mlTimer As Long

Private Sub Class_Terminate()

On Error Resume Next

If mlTimer <> 0 Then CloseHandle mlTimer
End Sub

' ÑÓÊ±
Private Sub DelayTime(ms As Long)
    On Error GoTo ErrHandler
    Dim ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    
    mlTimer = CreateWaitableTimer(0, True, App.EXEName & "Timer" & Format$(Now(), "NNSS"))
    
    If Err.LastDllError <> ERROR_ALREADY_EXISTS Then
    ft.dwLowDateTime = -1
    ft.dwHighDateTime = -1
    lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, 0)
    End If
    
    dblDelay = CDbl(ms) * 10000#
    ft.dwHighDateTime = -CLng(dblDelay / UNITS) - 1
    dblDelayLow = -UNITS * (dblDelay / UNITS - Fix(CStr(dblDelay / UNITS)))
    
    If dblDelayLow < MAX_LONG Then dblDelayLow = UNITS + dblDelayLow
    
    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, False)
    
    Do
    lBusy = MsgWaitForMultipleObjects(1, mlTimer, False, INFINITE, QS_ALLINPUT&)
    
    DoEvents
    Loop Until lBusy = WAIT_OBJECT_0
    
    CloseHandle mlTimer
    mlTimer = 0
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsWaitableTimer.Wait]" & Err.Description
End Sub

' ·¶Î§ÄÚËæ»úÑÓ³Ù
Public Sub Delay(min As Long, Optional max As Long = -2147483648#)
    Randomize
    If max = -2147483648# Then
        max = min
    End If
    If min < 10 Then
        min = 10
        max = 10
    End If
    DelayTime (Int((min - max + 1) * Rnd() + min))
End Sub
