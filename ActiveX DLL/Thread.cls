VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Thread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#If False Then
Private ThreadPriorityIdle, ThreadPriorityLowest, ThreadPriorityBelowNormal, ThreadPriorityNormal, ThreadPriorityAboveNormal, ThreadPriorityHighest, ThreadPriorityTimeCritical
#End If
Private Const THREAD_PRIORITY_IDLE As Long = -15
Private Const THREAD_PRIORITY_LOWEST As Long = -2
Private Const THREAD_PRIORITY_BELOW_NORMAL As Long = -1
Private Const THREAD_PRIORITY_NORMAL As Long = 0
Private Const THREAD_PRIORITY_ABOVE_NORMAL As Long = 1
Private Const THREAD_PRIORITY_HIGHEST As Long = 2
Private Const THREAD_PRIORITY_TIME_CRITICAL As Long = 15
Public Enum ThreadPriorityConstants
ThreadPriorityIdle = THREAD_PRIORITY_IDLE
ThreadPriorityLowest = THREAD_PRIORITY_LOWEST
ThreadPriorityBelowNormal = THREAD_PRIORITY_BELOW_NORMAL
ThreadPriorityNormal = THREAD_PRIORITY_NORMAL
ThreadPriorityAboveNormal = THREAD_PRIORITY_ABOVE_NORMAL
ThreadPriorityHighest = THREAD_PRIORITY_HIGHEST
ThreadPriorityTimeCritical = THREAD_PRIORITY_TIME_CRITICAL
End Enum
Public Event SuspendedChanged()
Attribute SuspendedChanged.VB_Description = "Occurs when the suspended state has changed."
Public Event PriorityChanged(ByVal OldPriority As ThreadPriorityConstants, ByVal NewPriority As ThreadPriorityConstants)
Attribute PriorityChanged.VB_Description = "Occurs when the priority has changed."
Implements IThreadStatusCallback
Private TI As TTHREADINFO, IID_IUnknown As ThreadAPI.CLSID
Private PropThreadOwner As IThread, PropThreadHandle As Long, PropThreadID As Long, PropThreadData As ThreadData
Private PropSuspended As Boolean

Private Sub IThreadStatusCallback_Raise(Optional ByRef Argument1 As Variant, Optional ByRef Argument2 As Variant)
If Not PropThreadOwner Is Nothing Then PropThreadOwner.StatusCallback TI.Key, Argument1, Argument2
End Sub

Private Sub Class_Initialize()
With IID_IUnknown
.Data4(0) = &HC0
.Data4(7) = &H46
End With
End Sub

Private Sub Class_Terminate()
If PropThreadHandle <> 0 Then Me.Terminate
End Sub

Public Sub DebugBackgroundProcedure(ByVal StatusCallback As IThreadStatusCallback, ByVal Data As ThreadData)
Attribute DebugBackgroundProcedure.VB_Description = "Internal method when a thread's background procedure is called in debug mode."
Attribute DebugBackgroundProcedure.VB_MemberFlags = "40"
Me.FBackgroundProcedure StatusCallback, Data
End Sub

Friend Sub FBackgroundProcedure(ByVal StatusCallback As IThreadStatusCallback, ByVal Data As ThreadData)
If Not PropThreadOwner Is Nothing Then PropThreadOwner.BackgroundProcedure TI.Key, StatusCallback, Data
End Sub

Friend Sub FComplete()
Dim Owner As IThread, Key As String, Data As ThreadData
Set Owner = PropThreadOwner
Key = TI.Key
Set Data = PropThreadData
Call Clear
If Not Owner Is Nothing Then Owner.Complete Key, Data
End Sub

Private Sub Clear()
With TI
Set .ShadowThread = Nothing
Set PropThreadOwner = Nothing
.Key = vbNullString
If PropThreadHandle <> 0 Then
    CloseHandle PropThreadHandle
    PropThreadHandle = 0
End If
PropThreadID = 0
Set PropThreadData = Nothing
PropSuspended = False
End With
End Sub

Public Function Create(ByVal Owner As IThread, Optional ByVal Key As String) As Boolean
Attribute Create.VB_Description = "Creates a new thread."
If PropThreadHandle <> 0 Then Err.Raise Number:=5, Description:="Thread already created"
If Owner Is Nothing Then Err.Raise Number:=380, Description:="Required argument is missing"
With TI
On Error Resume Next
Err.Raise 5
LSet .CLASSID = [_TA_OLE32].CLSIDFromProgID(Err.Source & ".Thread")
On Error GoTo 0
.lpStream = [_TA_OLE32].CoMarshalInterThreadInterfaceInStream(IID_IUnknown, Me)
.hEvent = CreateEvent(ByVal 0&, 1, 0, 0)
If .DebugMode = False Then Set .ShadowThread = Me
Set PropThreadOwner = Owner
Set PropThreadData = New ThreadData
PropThreadData.DebugMode = .DebugMode
.lpStreamData = [_TA_OLE32].CoMarshalInterThreadInterfaceInStream(IID_IUnknown, PropThreadData)
.Key = Key
PropThreadHandle = CreateThread(0&, 0, AddressOf ThreadBase.ThreadProc, ByVal VarPtr(TI), 0, PropThreadID)
PropSuspended = False
If PropThreadHandle <> 0 Then
    WaitForSingleObject .hEvent, INFINITE
    TimerCount = TimerCount + 1
    ReDim Preserve TMRI(1 To TimerCount) As TTIMERINFO
    With TMRI(TimerCount)
    Set .ShadowThread = Me
    .hThread = PropThreadHandle
    .ID = SetTimer(0, 0, 200, AddressOf ThreadBase.TimerProc)
    End With
Else
    Dim pStream As IUnknown
    If .lpStream <> 0 Then
        CopyMemory pStream, .lpStream, 4
        Set pStream = Nothing
    End If
    If .lpStreamData <> 0 Then
        CopyMemory pStream, .lpStreamData, 4
        Set pStream = Nothing
    End If
End If
If .hEvent <> 0 Then
    CloseHandle .hEvent
    .hEvent = 0
End If
End With
Create = CBool(PropThreadHandle <> 0)
End Function

Public Sub Cancel()
Attribute Cancel.VB_Description = "Sends a cancel request to the current thread."
If PropThreadHandle = 0 Then Err.Raise Number:=5, Description:="Thread must be created before it can be used"
If Not PropThreadData Is Nothing Then PropThreadData.CancellationPending = True
End Sub

Public Function Terminate(Optional ByVal Wait As Boolean) As Boolean
Attribute Terminate.VB_Description = "Terminates the current thread."
If PropThreadHandle = 0 Then Err.Raise Number:=5, Description:="Thread must be created before it can be used"
Dim dwExitCode As Long
If GetExitCodeThread(PropThreadHandle, dwExitCode) <> 0 Then
    If Wait = False Then
        If TerminateThread(PropThreadHandle, dwExitCode) <> 0 Then Call Clear
    Else
        While TerminateThread(PropThreadHandle, dwExitCode) = 0
            Sleep 200
        Wend
        Call Clear
    End If
End If
Terminate = CBool(PropThreadHandle = 0)
End Function

Public Property Get hThread() As Long
Attribute hThread.VB_Description = "Returns a handle to a thread."
Attribute hThread.VB_UserMemId = 0
hThread = PropThreadHandle
End Property

Public Property Get ThreadID() As Long
Attribute ThreadID.VB_Description = "Returns an internal identification of a thread."
ThreadID = PropThreadID
End Property

Public Property Get Suspended() As Boolean
Attribute Suspended.VB_Description = "Returns/sets a value that determines if a thread is suspended."
If PropThreadHandle = 0 Then Err.Raise Number:=5, Description:="Thread must be created before it can be used"
Suspended = PropSuspended
End Property

Public Property Let Suspended(ByVal Value As Boolean)
If PropThreadHandle = 0 Then Err.Raise Number:=5, Description:="Thread must be created before it can be used"
If PropSuspended = Value Then Exit Property
Dim RetVal As Long, Changed As Boolean
If Value = True Then
    If PropSuspended = False Then
        RetVal = SuspendThread(PropThreadHandle)
        Changed = CBool(RetVal <> &HFFFFFFFF)
    End If
Else
    If PropSuspended = True Then
        RetVal = ResumeThread(PropThreadHandle)
        Changed = CBool(RetVal <> &HFFFFFFFF)
    End If
End If
PropSuspended = Value
If Changed = True Then RaiseEvent SuspendedChanged
End Property

Public Property Get Priority() As ThreadPriorityConstants
Attribute Priority.VB_Description = "Returns/sets the priority of a thread."
If PropThreadHandle = 0 Then Err.Raise Number:=5, Description:="Thread must be created before it can be used"
Priority = GetThreadPriority(PropThreadHandle)
End Property

Public Property Let Priority(ByVal Value As ThreadPriorityConstants)
If PropThreadHandle = 0 Then Err.Raise Number:=5, Description:="Thread must be created before it can be used"
Select Case Value
    Case ThreadPriorityIdle, ThreadPriorityLowest, ThreadPriorityBelowNormal, ThreadPriorityNormal, ThreadPriorityAboveNormal, ThreadPriorityHighest, ThreadPriorityTimeCritical
        Dim OldPriority As ThreadPriorityConstants
        OldPriority = GetThreadPriority(PropThreadHandle)
        If OldPriority = Value Then Exit Property
        If SetThreadPriority(PropThreadHandle, Value) <> 0 Then RaiseEvent PriorityChanged(OldPriority, Value)
    Case Else
        Err.Raise 380
End Select
End Property

Public Property Get DebugMode() As Boolean
Attribute DebugMode.VB_Description = "Returns/sets a value that indicates if the background procedure will run on the main thread. This does allow to safely set breakpoints and perform debugging."
DebugMode = TI.DebugMode
End Property

Public Property Let DebugMode(ByVal Value As Boolean)
If PropThreadHandle <> 0 Then Err.Raise Number:=383, Description:="DebugMode property is read-only when thread is created"
TI.DebugMode = Value
End Property
