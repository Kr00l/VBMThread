VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub BackgroundProcedure(ByVal Key As String, ByVal StatusCallback As IThreadStatusCallback, ByVal Data As ThreadData)
Attribute BackgroundProcedure.VB_Description = "Interface method when a thread's background procedure is called."
End Sub

Public Sub Complete(ByVal Key As String, ByVal Data As ThreadData)
Attribute Complete.VB_Description = "Interface method when a thread's background procedure has been completed or canceled."
End Sub

Public Sub StatusCallback(ByVal Key As String, ByRef Argument1 As Variant, ByRef Argument2 As Variant)
Attribute StatusCallback.VB_Description = "Interface method when a thread requests an synchronous status callback. The background thread will be suspended, the method request is marshaled to the main thread."
End Sub
