VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_BetterAccess_OpenAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2023 Thomas Moeller                                                       #
'# MIT License  => https://github.com/team-moeller/better-access-open-ai/blob/main/LICENSE #
'# Version 0.91.04  published: 06.08.2023                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_API_Key As String


' ### Properties

Public Property Let API_Key(ByVal This_API_Key As String)
    m_API_Key = This_API_Key
End Property
Public Property Get API_Key() As String
    API_Key = m_API_Key
End Property


'### Instancing

Private Sub Class_Initialize()
    
    'Me.API_Key = "INSERT YOUR OPEN AI API-KEY HERE"
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub CreateCompletion(fSuccess As Boolean, Response As String)

    'Declarations
    Dim request As Object
    Dim RequestBody As String
    
    Set request = CreateObject("MSXML2.ServerXMLHTTP")
    
    request.Open "POST", "https://api.openai.com/v1/chat/completions", False
    request.setRequestHeader "Content-Type", "application/json"
    request.setRequestHeader "Authorization", "Bearer " & Me.API_Key
    
    RequestBody = "{"
    RequestBody = RequestBody & """model"": ""gpt-3.5-turbo"","
    RequestBody = RequestBody & """max_tokens"": " & "500"
    RequestBody = RequestBody & "," & """messages"": " & "[{""role"": ""user"", ""content"": """ & "Please introduce yourself." & """}]"
    RequestBody = RequestBody & "}"
    
    request.send RequestBody

    If request.Status = 200 Then
        fSuccess = True
    Else
        fSuccess = False
    End If
    
    Response = request.responseText

End Sub

    
'### Private Methods

Private Sub DoSomethingPrivate()

    'Declarations

End Sub


