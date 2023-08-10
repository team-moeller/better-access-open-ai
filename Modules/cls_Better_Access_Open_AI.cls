VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Better_Access_Open_AI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'###########################################################################################
'# Copyright (c) 2023 Thomas Moeller                                                       #
'# MIT License  => https://github.com/team-moeller/better-access-open-ai/blob/main/LICENSE #
'# Version 0.93.04  published: 10.08.2023                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Members

Private m_API_Key As String
Private m_Max_Tokens As Long


' ### Properties

Public Property Let API_Key(ByVal This_API_Key As String)
    m_API_Key = This_API_Key
End Property
Public Property Get API_Key() As String
    API_Key = m_API_Key
End Property

Public Property Let Max_Tokens(ByVal This_Max_Tokens As Long)
    m_Max_Tokens = This_Max_Tokens
End Property
Public Property Get Max_Tokens() As Long
    Max_Tokens = m_Max_Tokens
End Property


'### Instancing

Private Sub Class_Initialize()
    
    'Me.API_Key = "INSERT YOUR OPEN AI API-KEY HERE"
    Me.Max_Tokens = 500
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub CreateCompletion(ByRef fSuccess As Boolean, ByRef Response As String)

    'Declarations
    Dim Request As Object
    Dim RequestBody As String
    
    Set Request = CreateObject("MSXML2.ServerXMLHTTP")
    
    Request.Open "POST", GetApiEndpoint("gpt-3.5-turbo"), False
    Request.setRequestHeader "Content-Type", "application/json"
    Request.setRequestHeader "Authorization", "Bearer " & Me.API_Key
    
    RequestBody = "{"
    RequestBody = RequestBody & """model"": ""gpt-3.5-turbo"","
    RequestBody = RequestBody & """max_tokens"": " & Me.Max_Tokens
    RequestBody = RequestBody & "," & """messages"": " & "[{""role"": ""user"", ""content"": """ & "Please introduce yourself." & """}]"
    RequestBody = RequestBody & "}"
    
    Request.send RequestBody

    If Request.Status = 200 Then
        fSuccess = True
    Else
        fSuccess = False
    End If
    
    Response = Request.responseText

End Sub

    
'### Private Methods

Private Function GetApiEndpoint(ByVal Model As String) As String

    'Declarations
    Dim ApiEndpoint As String
    
    Select Case Model
        Case "gpt-4"
            ApiEndpoint = "https://api.openai.com/v1/chat/completions"
        Case "gpt-3.5-turbo"
            ApiEndpoint = "https://api.openai.com/v1/chat/completions"
        Case "text-davinci-003"
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case "text-davinci-002"
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case "davinci"
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case "curie"
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case "babbage"
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case "ada"
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case Else
            ApiEndpoint = vbNullString
    End Select
    
    GetApiEndpoint = ApiEndpoint

End Function


