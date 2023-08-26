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
'# Version 0.96.02  published: 26.08.2023                                                  #
'###########################################################################################

Option Compare Database
Option Explicit


'### Enums

Public Enum oaiModel
    gpt4 = 1
    gpt35turbo = 2
    TextDavinci3 = 3
    TextDavinci2 = 4
    davinci = 5
    curie = 6
    babbage = 7
    ada = 8
End Enum


'### Members

Private m_API_Key As String
Private m_Max_Tokens As Long
Private m_Model As oaiModel


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

Public Property Let Model(ByVal This_Model As oaiModel)
    m_Model = This_Model
End Property
Public Property Get Model() As oaiModel
    Model = m_Model
End Property


'### Instancing

Private Sub Class_Initialize()
    
    'Me.API_Key = "INSERT YOUR OPEN AI API-KEY HERE"
    Me.Max_Tokens = 500
    Me.Model = oaiModel.gpt35turbo
    
End Sub

Private Sub Class_Terminate()
    ' Do Nothing
End Sub


'### Public Methods

Public Sub CreateCompletion(ByRef fSuccess As Boolean, ByRef Response As String)

    'Declarations
    Dim Request As Object
    
    Set Request = CreateObject("MSXML2.ServerXMLHTTP")
    
    Request.Open "POST", getApiEndpoint(), False
    Request.setRequestHeader "Content-Type", "application/json"
    Request.setRequestHeader "Authorization", "Bearer " & Me.API_Key
    
    Request.send getRequestBody

    If Request.Status = 200 Then
        fSuccess = True
    Else
        fSuccess = False
    End If
    
    Response = getResponseMessage(Request.responseText)

End Sub

    
'### Private Methods

Private Function getApiEndpoint() As String

    'Declarations
    Dim ApiEndpoint As String
    
    Select Case Me.Model
        Case oaiModel.gpt4
            ApiEndpoint = "https://api.openai.com/v1/chat/completions"
        Case oaiModel.gpt35turbo
            ApiEndpoint = "https://api.openai.com/v1/chat/completions"
        Case oaiModel.TextDavinci3
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case oaiModel.TextDavinci2
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case oaiModel.davinci
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case oaiModel.curie
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case oaiModel.babbage
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case oaiModel.ada
            ApiEndpoint = "https://api.openai.com/v1/completions"
        Case Else
            ApiEndpoint = vbNullString
    End Select
    
    getApiEndpoint = ApiEndpoint

End Function

Private Function getRequestBody()

    'Declarations
    Dim RequestBody As String
    
    Select Case Me.Model
        Case oaiModel.gpt35turbo, oaiModel.gpt4
            RequestBody = "{"
            RequestBody = RequestBody & """model"": """ & getModelAsText & ""","
            RequestBody = RequestBody & """max_tokens"": " & Me.Max_Tokens
            RequestBody = RequestBody & "," & """messages"": " & "[{""role"": ""user"", ""content"": """ & "Please introduce yourself." & """}]"
            RequestBody = RequestBody & "}"
        Case Else
            RequestBody = "{"
            RequestBody = RequestBody & """model"": """ & getModelAsText & ""","
            RequestBody = RequestBody & """max_tokens"": " & Me.Max_Tokens
            RequestBody = RequestBody & "," & """prompt"": " & """Please introduce yourself." & """"
            RequestBody = RequestBody & "}"
    End Select
    getRequestBody = RequestBody

End Function

Private Function getModelAsText() As String

    Select Case Me.Model
        Case oaiModel.ada
            getModelAsText = "ada"
        Case oaiModel.babbage
            getModelAsText = "babbage"
        Case oaiModel.curie
            getModelAsText = "curie"
        Case oaiModel.davinci
            getModelAsText = "davinci"
        Case oaiModel.gpt35turbo
            getModelAsText = "gpt-3.5-turbo"
        Case oaiModel.gpt4
            getModelAsText = "gpt-4"
        Case oaiModel.TextDavinci2
            getModelAsText = "text-davinci-002"
        Case oaiModel.TextDavinci3
            getModelAsText = "text-davinci-003"
        Case Else
            getModelAsText = "gpt-3.5-turbo"
    End Select

End Function

Private Function getResponseMessage(ByVal responseText As String) As String

    'Declarations
    Dim Message As String
    
    Select Case Me.Model
        Case oaiModel.gpt35turbo, oaiModel.gpt4
            Message = Right$(responseText, Len(responseText) - InStrRev(responseText, Chr(34) & "content" & Chr(34) & ": " & Chr(34)) - 11)
            Message = Left$(Message, InStrRev(Message, Chr(34) & Chr(10) & "      }," & Chr(10) & "      ") - 1)
        Case Else
            Message = Right$(responseText, Len(responseText) - InStrRev(responseText, "[" & Chr(10) & "    {" & Chr(10) & "      " & Chr(34) & "text" & Chr(34) & ": " & Chr(34)) - 22)
            Message = Left$(Message, InStrRev(Message, Chr(34) & "," & Chr(10) & "      " & Chr(34) & "index") - 1)
            Message = Replace(Message, "\n", Chr(10))
            If Left$(Message, 2) = Chr(10) & Chr(10) Then
                Message = Right$(Message, Len(Message) - 2)
            End If
    End Select

    getResponseMessage = Message
    
End Function

