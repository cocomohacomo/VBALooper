Attribute VB_Name = "Looper"
Option Explicit

Private MainLooper As VBALooper.CallBackLooper

Public Sub Refresh()
If MainLooper Is Nothing Then Exit Sub
MainLooper.Refresh
End Sub

Public Sub AddHandler(ByRef Handler As IHandler)
If MainLooper Is Nothing Then Set MainLooper = New VBALooper.CallBackLooper
MainLooper.AddHandler Handler
MainLooper.StartCallback
End Sub

Public Sub RemoveHandler(ByRef Handler As IHandler)
If MainLooper Is Nothing Then Exit Sub
If MainLooper.RemoveHandler(Handler) = 0 Then MainLooper.StopCallBack
End Sub

