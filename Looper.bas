Attribute VB_Name = "Looper"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)

Private LoopActive As Boolean 'Default:False
Private Handlers As Object

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
Dim Handles As Variant
Dim i As Long
Do
    If Not LoopaActive Then Exit Do
    If Handlers Is Nothing Then Exit Do
    If Handlers.Count = 0 Then Exit Do
    Handles = Handlers.Keys
    For i = 0 To UBound(Handles)
        Handles(i).CallBack
    Next
    Sleep 1&
    VBA.DoEvents
Loop
LoopActive = False
End Sub
