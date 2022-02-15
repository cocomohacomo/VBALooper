Attribute VB_Name = "Looper"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)

Private Handlers As Object
Private LoopActive As Boolean

Public Sub Start(ByVal HandlersPtr As String)
If CLngPtr(HandlersPtr) = ObjPtr(Handlers) Then: LoopActive = True: Call Looper
End Sub

Public Sub Refresh()
If Handlers Is Nothing Then Exit Sub
LoopActive = False
Handlers.RemoveAll
Set Handlers = Nothing
End Sub

Public Sub AddHandler(ByRef Handler As IHandler)
If Handlers Is Nothing Then Set Handlers = CreateObject("Scripting.Dictionary")
If Not Handlers.Exists(Handler) Then Handlers.Add Handler, vbNullString
If Not LoopActive Then Application.OnTime Now(), "'Start """ & CStr(ObjPtr(Handlers)) & """ '"
End Sub

Public Sub RemoveHandler(ByRef Handler As IHandler)
If Handlers Is Nothing Then Exit Sub
If Handlers.Exists(Handler) Then Handlers.Remove Handler
If Handlers.Count = 0 Then LoopActive = False
End Sub

Private Sub Looper()
Dim Handler As IHandler
Do
    For Each Handler In Handlers
        Handler.CallBack
        If Not LoopActive Then Exit Do
        If Handlers Is Nothing Then Exit Do
        If Handlers.Count = 0 Then Exit Do
        If Handler Is Nothing Then Exit For
        VBA.DoEvents
        Sleep 1&
    Next
    If Not LoopActive Then Exit Do
    If Handlers Is Nothing Then Exit Do
    If Handlers.Count = 0 Then Exit Do
Loop
LoopActive = False
End Sub
