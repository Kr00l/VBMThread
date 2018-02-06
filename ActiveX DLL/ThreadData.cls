VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private PropCancellationPending As Boolean
Private PropCanceled As Boolean
Private PropDebugMode As Boolean
Private PropTag As Variant

Public Property Get CancellationPending() As Boolean
Attribute CancellationPending.VB_Description = "Returns a value that determines if a cancellation request is pending."
CancellationPending = PropCancellationPending
End Property

Friend Property Let CancellationPending(ByVal Value As Boolean)
PropCancellationPending = Value
End Property

Public Property Get Canceled() As Boolean
Attribute Canceled.VB_Description = "Returns/sets a value that determines if the background procedure ran to completion or if it was canceled."
Canceled = PropCanceled
End Property

Public Property Let Canceled(ByVal Value As Boolean)
PropCanceled = Value
End Property

Public Property Get DebugMode() As Boolean
Attribute DebugMode.VB_Description = "Returns a value that indicates if the background procedure is running on the main thread. This does allow to safely set breakpoints and perform debugging."
DebugMode = PropDebugMode
End Property

Friend Property Let DebugMode(ByVal Value As Boolean)
PropDebugMode = Value
End Property

Public Property Get Tag() As Variant
Attribute Tag.VB_Description = "Stores any extra data needed for your program."
If IsEmpty(PropTag) Then
    Tag = vbNullString
ElseIf IsObject(PropTag) Then
    Set Tag = PropTag
Else
    Tag = PropTag
End If
End Property

Public Property Let Tag(ByVal Value As Variant)
PropTag = Value
End Property

Public Property Set Tag(ByVal Value As Variant)
Set PropTag = Value
End Property
