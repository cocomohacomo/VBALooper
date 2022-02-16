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
        Select Case True
            Case Handlers Is Nothing: Exit Do
            Case Handlers.Count = 0: Exit Do
            Case Handler Is Nothing: Exit For
            Case Not LoopActive: Exit Do
            Case Else
                VBA.DoEvents
                Sleep 1&
        End Select
    Next
    Select Case True
        Case Not LoopActive: Exit Do
        Case Handlers Is Nothing: Exit Do
        Case Handlers.Count = 0: Exit Do
    End Select
Loop
LoopActive = False
End Sub
