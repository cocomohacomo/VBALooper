Attribute VB_Name = "Looper"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)

Private LoopaActive As Boolean 'Default:False
Private Handlers As Object

Public Sub Start(ByVal HandlersPtr As String)
If CLngPtr(HandlersPtr) = ObjPtr(Handlers) Then: LoopaActive = True: Call Looper
End Sub

Public Sub Refresh()
If Handlers Is Nothing Then Exit Sub
LoopaActive = False
Handlers.RemoveAll
Set Handlers = Nothing
End Sub

Public Sub AddHandler(ByRef Handler As IHandler)
If Handlers Is Nothing Then Set Handlers = CreateObject("Scripting.Dictionary")
If Not Handlers.Exists(Handler) Then Handlers.Add Handler, vbNullString
If Not LoopaActive Then Application.OnTime Now(), "'Start """ & CStr(ObjPtr(Handlers)) & """ '"
End Sub

Public Sub RemoveHandler(ByRef Handler As IHandler)
If Handlers Is Nothing Then Exit Sub
If Handlers.Exists(Handler) Then Handlers.Remove Handler
If Handlers.Count = 0 Then LoopaActive = False
End Sub

Private Sub Looper()
Dim Handler As IHandler
Do
    For Each Handler In Handlers
        Handler.CallBack
        If Handlers Is Nothing Then Exit Do
        If Handlers.Count = 0 Then Exit Do
        If Not LoopaActive Then Exit Do
        If Handler Is Nothing Then Exit For
    Next
    If Handlers Is Nothing Then Exit Do
    If Handlers.Count = 0 Then Exit Do
    If Not LoopaActive Then Exit Do
    Sleep CLng(VBA.DoEvents + 1)
Loop
LoopaActive = False
End Sub
