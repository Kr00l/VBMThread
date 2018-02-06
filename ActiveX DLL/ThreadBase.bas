Attribute VB_Name = "ThreadBase"
Option Explicit
Public Type TTHREADINFO
CLASSID As CLSID
lpStream As Long
hEvent As Long
lpStreamData As Long
DebugMode As Boolean
ShadowThread As Thread
Key As String
End Type
Public Type TTIMERINFO
ID As Long
hThread As Long
ShadowThread As Thread
End Type
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function CreateThread Lib "kernel32" (ByVal lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByRef lpParameter As Any, ByVal dwCreationFlags As Long, ByRef lpThreadID As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateEvent Lib "kernel32" Alias "CreateEventW" (ByRef lpEventAttributes As Any, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Public Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Const INFINITE As Long = (-1)
Public Const WM_TIMER As Long = &H113
Public Const WAIT_OBJECT_0 As Long = &H0
Public TimerCount As Long
Public TMRI() As TTIMERINFO

Public Function ThreadProc(ByRef TI As TTHREADINFO) As Long
Dim IUnk As IUnknown, IID_IUnknown As ThreadAPI.CLSID
[_TA_OLE32].CoInitialize 0
With IID_IUnknown
.Data4(0) = &HC0
.Data4(7) = &H46
End With
With TI
[_TA_OLE32].CoCreateInstance .CLASSID, Nothing, ThreadAPI.CLSCTX_INPROC_SERVER, IID_IUnknown, IUnk
SetEvent .hEvent
Dim pStream As IUnknown, pStreamData As IUnknown
Set pStream = [_TA_OLE32].CoGetInterfaceAndReleaseStream(.lpStream, IID_IUnknown)
Set pStreamData = [_TA_OLE32].CoGetInterfaceAndReleaseStream(.lpStreamData, IID_IUnknown)
If .DebugMode = False Then
    .ShadowThread.FBackgroundProcedure pStream, pStreamData
Else
    Dim This As Thread
    Set This = pStream
    This.DebugBackgroundProcedure pStream, pStreamData
End If
End With
Set IUnk = Nothing
Set pStream = Nothing
Set pStreamData = Nothing
[_TA_OLE32].CoUninitialize
End Function

Public Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
If uMsg = WM_TIMER And TimerCount > 0 Then
    Dim i As Long, Index As Long
    For i = 1 To TimerCount
        If TMRI(i).ID = wParam Then
            Index = i
            Exit For
        End If
    Next i
    If Index > 0 Then
        Dim This As Thread
        With TMRI(Index)
        If WaitForSingleObject(.hThread, 0) = WAIT_OBJECT_0 Then
            KillTimer 0, .ID
            Set This = .ShadowThread
        End If
        End With
        If Not This Is Nothing Then
            Dim j As Long
            For j = Index To TimerCount - 1
                LSet TMRI(j) = TMRI(j + 1)
            Next j
            TimerCount = TimerCount - 1
            If TimerCount > 0 Then
                ReDim Preserve TMRI(1 To TimerCount) As TTIMERINFO
            Else
                Erase TMRI()
            End If
            This.FComplete
        End If
    End If
End If
End Sub
