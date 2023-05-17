VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Project: OpenAI VBA Framework
' Class:  clsOpenAI
' Description: Main class that controls the framework
'
' Author: Zaid Qureshi
' GitHub: https://github.com/zq99
'
' Classes / Modules in the Framework:
' - clsOpenAI
' - clsOpenAILogger
' - clsOpenAIMessage
' - clsOpenAIMessages
' - clsOpenAIRequest
' - clsOpenAIResponse
' - IOpenAINameProvider
'
' - mdOpenAI_Tests
' - mdOpenAI_Examples
'
' This work is licensed under the MIT License. The full license text
' can be found in the LICENSE file in the root of this repository.
'
' Copyright (c) 2023 Zaid Qureshi
'-----------------------------------------------------------------------------

Option Explicit

Implements IOpenAINameProvider

Private mstrAPI_KEY As String
Private mobjHttpRequest As Object
Private mobjLogger As clsOpenAILogger
Private mobjRequest As clsOpenAIRequest

'Open AI defined constants
Private Const API_ENDPOINT_CHAT As String = "https://api.openai.com/v1/chat/completions"
Private Const API_ENDPOINT_COMPLETIONS As String = "https://api.openai.com/v1/completions"

'More models can be found here: https://platform.openai.com/docs/models/overview
Private Const DEFAULT_CHAT_MODEL As String = "gpt-3.5-turbo"
Private Const DEFAULT_TEXT_COMPLETION_MODEL As String = "text-davinci-003"

Private Const DEFAULT_CHAT_TOKENS_COUNT As Integer = 512
Private Const DEFAULT_TEXT_TOKENS_COUNT As Integer = 1024

Private Const UNASSIGNED_VALUE As Integer = -1
Private Const MESSAGE_INVALID_API_KEY As String = "An OpenAI API key is either invalid or has not been specified!"
Private Const HTTP_STATUS_OK As Long = 200 ' OK
Private Const HTTP_REQUEST_COMPLETED As Integer = 4


Private Function IOpenAINameProvider_GetClassName() As String
    IOpenAINameProvider_GetClassName = "clsOpenAI"
End Function

Private Function IOpenAINameProvider_ToString() As String
    IOpenAINameProvider_ToString = "Key=" & Me.API_KEY
End Function

Public Property Let API_KEY(ByVal value As String)
    mstrAPI_KEY = value
End Property

Public Property Get API_KEY() As String
    API_KEY = mstrAPI_KEY
End Property

Public Property Let Model(ByVal value As String)
     mobjRequest.Model = value
End Property

Public Property Let MaxTokens(ByVal value As Long)
    mobjRequest.MaxTokens = value
End Property

Public Property Let TopP(ByVal value As Double)
    mobjRequest.TopP = value
End Property

Public Property Let Temperature(ByVal value As Double)
    If (value < 0) Or (value > 1) Then
        Call mobjLogger.LogCriticalMessage("Temperature setting must be between 0 and 1!", blnAddModuleName:=False)
    End If
    mobjRequest.Temperature = value
End Property

Public Property Let FrequencyPenalty(ByVal value As Double)
    mobjRequest.FrequencyPenalty = value
End Property

Public Property Let PresencePenalty(ByVal value As Double)
    mobjRequest.PresencePenalty = value
End Property


Public Sub IsLogOutputRequired(ByVal value As Boolean)
'Purpose: Calling routines can switch off messages in this framework from appearing in the Immediate window

    If Not mobjLogger Is Nothing Then
        mobjLogger.IsMessageRequired = value
    End If

End Sub


Public Sub Log(ByVal strMessage As String)
'Purpose: Easy routine to log messages

    If Not mobjLogger Is Nothing Then
        mobjLogger.PrintMessage strMessage
    End If
End Sub


Private Function GetResponseFromAPI(ByVal strRequestJson As String, ByVal strEndPoint As String) As clsOpenAIResponse
'Purpose: This handles the request to OpenAI's API URL

    Dim strResponseJson As String
    Dim oResponse As clsOpenAIResponse
    
On Error GoTo ERR_HANDLER:

    'default return value
    Set GetResponseFromAPI = Nothing
    
    If mobjHttpRequest Is Nothing Then
        GoTo EXIT_HERE
    End If
    
    'talk to OpenAI
    With mobjHttpRequest
        .Open "POST", strEndPoint, False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & mstrAPI_KEY
        .send (strRequestJson)
    End With
    
    ' completed
    Do While mobjHttpRequest.readyState <> HTTP_REQUEST_COMPLETED
        DoEvents
    Loop
    
    Log "Response code from OpenAI API is: " & mobjHttpRequest.Status

    If mobjHttpRequest.Status = HTTP_STATUS_OK Then
    
        'get the json result from the successful request
        strResponseJson = Trim(mobjHttpRequest.responseText)
        Log strResponseJson
        Set oResponse = New clsOpenAIResponse
        
        'format the json result according to which api endpoint used
        If strEndPoint = API_ENDPOINT_CHAT Then
            
            'ChatGPT and GPT4
            oResponse.ParseChatJSON (strResponseJson)
            Set GetResponseFromAPI = oResponse
        
        ElseIf strEndPoint = API_ENDPOINT_COMPLETIONS Then
            
            'GPT3 and earlier
            oResponse.ParseTextCompletionJSON (strResponseJson)
            Set GetResponseFromAPI = oResponse
        
        End If
    End If
    
EXIT_HERE:
    Set oResponse = Nothing
    Exit Function
    
ERR_HANDLER:
    mobjLogger.LogVBAError Err
    GoTo EXIT_HERE
End Function


Private Function IsAPIKeyValid() As Boolean
'Purpose: Check a valid API key has been assigned

    IsAPIKeyValid = IIf(Trim(mstrAPI_KEY) = Empty, False, True)

End Function


Public Function ChatCompletion(ByVal oMessages As clsOpenAIMessages) As clsOpenAIResponse
'Purpose: This is for OpenAI's ChatGPT and GPT4 API

    Set ChatCompletion = Nothing

    If Not IsAPIKeyValid Then
        mobjLogger.LogCriticalMessage MESSAGE_INVALID_API_KEY, True
        Exit Function
    End If

    If mobjHttpRequest Is Nothing Or oMessages Is Nothing Then
        Exit Function
    End If
    
    Set mobjRequest.messages = oMessages
    
    If mobjRequest.Model = Empty Then
        mobjRequest.Model = DEFAULT_CHAT_MODEL
    End If
    
    If mobjRequest.MaxTokens = UNASSIGNED_VALUE Then
        mobjRequest.MaxTokens = DEFAULT_CHAT_TOKENS_COUNT
    End If
    
    Log mobjRequest.GetChatSendToAPIJsonString
    
    Set ChatCompletion = GetResponseFromAPI(mobjRequest.GetChatSendToAPIJsonString, API_ENDPOINT_CHAT)
    
End Function


Public Function TextCompletion(ByVal strPrompt As String) As clsOpenAIResponse
'Purpose: This is for OpenAI's Text Completion API

    Set TextCompletion = Nothing

    If Not IsAPIKeyValid Then
        mobjLogger.LogCriticalMessage MESSAGE_INVALID_API_KEY, True
        Exit Function
    End If
    
    If mobjHttpRequest Is Nothing Or strPrompt = Empty Then
        Exit Function
    End If
    
    mobjRequest.prompt = strPrompt
    
    If mobjRequest.Model = Empty Then
        mobjRequest.Model = DEFAULT_TEXT_COMPLETION_MODEL
    End If
    
    If mobjRequest.MaxTokens = UNASSIGNED_VALUE Then
        mobjRequest.MaxTokens = DEFAULT_TEXT_TOKENS_COUNT
    End If
    
    Log mobjRequest.GetTextCompletionSendToAPIJsonString
    
    Set TextCompletion = GetResponseFromAPI(mobjRequest.GetTextCompletionSendToAPIJsonString, API_ENDPOINT_COMPLETIONS)

End Function


Private Sub Class_Initialize()
    
    Set mobjHttpRequest = CreateObject("MSXML2.XMLHTTP")
    Set mobjRequest = GetDefaultRequestSettings

    Set mobjLogger = New clsOpenAILogger
    mobjLogger.IsMessageRequired = False
    mobjLogger.SetClass Me
    
    mstrAPI_KEY = Empty

End Sub


Private Sub Class_Terminate()
    Set mobjHttpRequest = Nothing
    Set mobjLogger = Nothing
    Set mobjRequest = Nothing
End Sub


Private Function GetDefaultRequestSettings() As clsOpenAIRequest
'Purpose: These are initial settings for the OpenAI request

    Dim oRequest As clsOpenAIRequest
    Set oRequest = New clsOpenAIRequest
    
    With oRequest
        .Model = Empty
        .MaxTokens = UNASSIGNED_VALUE
        .TopP = 1
        .Temperature = 0.5
        .FrequencyPenalty = 0
        .PresencePenalty = 0
    End With
    Set GetDefaultRequestSettings = oRequest
    
    Set oRequest = Nothing
End Function


Public Sub ClearSettings()
'Purpose: Reset the settings if switching between endpoints

    Set mobjRequest = GetDefaultRequestSettings

End Sub

Public Function GetReadAPIKeyFromFolder(ByVal strfolderPath As String) As String
'Purpose: Allows retrieval of an API KEY saved in an external file (possibly stored on a user access drive)

    Dim intFileNumber As Integer
    Dim strAPIKey As String
    Dim strFilePath As String

    strFilePath = strfolderPath & "\apikey.txt" ' Construct the full file path

    intFileNumber = FreeFile ' Get the first available file number

    ' Open the file in Input mode. Trappable error occurs if file does not exist.
    Open strFilePath For Input As intFileNumber

        ' Read the contents of the file into the variable.
        Input #intFileNumber, strAPIKey

    ' Close the file.
    Close intFileNumber

    ' Return the API Key
    GetReadAPIKeyFromFolder = strAPIKey
End Function


