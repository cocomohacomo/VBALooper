Attribute VB_Name = "Looper"
Option Explicit

Private MainLooper As VBALooper.CallBackLooper

Public Sub Start(ByVal MainLoopPtr As String)
If CLngPtr(MainLoopPtr) <> ObjPtr(MainLooper) Then Exit Sub
MainLooper.StartCallback
End Sub

Public Sub Refresh()
If MainLooper Is Nothing Then Exit Sub
MainLooper.RefreshHandler
End Sub

Public Sub AddHandler(ByRef Handler As IHandler)
If MainLooper Is Nothing Then Set MainLooper = New VBALooper.CallBackLooper
MainLooper.AddHandler Handler
If Not MainLooper.LoopStatus Then Application.OnTime Now(), "'Start """ & CStr(ObjPtr(MainLooper)) & """ '"
End Sub

Public Sub RemoveHandler(ByRef Handler As IHandler)
If MainLooper Is Nothing Then Exit Sub
If MainLooper.RemoveHandler(Handler) = 0 Then MainLooper.StopCallBack
End Sub

Public Property Get LoopActive() As Boolean
LoopActive = MainLooper.LoopStatus
End Property

