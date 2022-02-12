VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallBackLooper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub Sleep _
    Lib "kernel32" ( _
        ByVal Milliseconds As Long)

Private Handlers As Object
Private LoopActive As Boolean

Public Property Get LoopStatus() As Boolean
LoopStatus = LoopActive
End Property

Public Function AddHandler(ByRef Handler As IHandler) As Long
If Not Handlers.Exists(Handler) Then Handlers.Add Handler, Handler
AddHandler = Handlers.Count
End Function

Public Function RemoveHandler(ByRef Handler As IHandler) As Long
If Handlers.Exists(Handler) Then Handlers.Remove Handler
RemoveHandler = Handlers.Count
End Function

Public Function ExistsHandler(ByRef Handler As IHandler) As Boolean
ExistsHandler = Handlers.Exists(Handler)
End Function

Public Sub RefreshHandler()
Handlers.RemoveAll
End Sub

Public Sub StartCallback()
Call Looper
End Sub

Public Sub StopCallBack()
LoopActive = False
End Sub

Private Sub Looper()
If LoopActive Then Exit Sub Else LoopActive = True
Dim Handler As IHandler
Do While LoopActive
    For Each Handler In Handlers
        Handler.CallBack
        VBA.DoEvents
        Sleep 1&
        If Handler Is Nothing Then Exit For
    Next
Loop
End Sub

Private Sub Class_Initialize()
Set Handlers = CreateObject("Scripting.Dictionary")
Handlers.CompareMode = 0& 'BinaryCompare
LoopActive = False
End Sub

Private Sub Class_Terminate()
Set Handlers = Nothing
End Sub
