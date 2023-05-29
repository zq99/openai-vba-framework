VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenAIRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Project: OpenAI VBA Framework
' Class:  clsOpenAIRequest
' Description: Builds the request json to send to OpenAI API
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
'-----------------------------------------------------------------------------

Option Explicit

Implements IOpenAINameProvider

Private mstrModel As String
Private mobjMessages As clsOpenAIMessages
Private mlngMaxTokens As Long
Private mdblTopP As Double
Private mdblTemperature As Double
Private mdblFrequencyPenalty As Double
Private mdlPresencePenalty As Double
Private mstrPrompt As String
Private mlngImageWidth As Long
Private mlngImageHeight As Long

'Resolve, Connect, Send and Receive
Private mlngTimeoutResolve As Long
Private mlngTimeoutConnect As Long
Private mlngTimeoutSend As Long
Private mlngTimeoutReceive As Long


Private Function IOpenAINameProvider_GetClassName() As String
    IOpenAINameProvider_GetClassName = "clsOpenAIRequest"
End Function


Private Function IOpenAINameProvider_ToString() As String
    IOpenAINameProvider_ToString = "chat json=" & Me.GetChatSendToAPIJsonString & " | " & "text json=" & Me.GetTextCompletionSendToAPIJsonString
End Function


Private Sub Class_Initialize()
    Set mobjMessages = New clsOpenAIMessages
End Sub


Private Sub Class_Terminate()
    Set mobjMessages = Nothing
End Sub


Public Property Let prompt(ByVal value As String)
    mstrPrompt = value
End Property

Public Property Get Model() As String
    Model = mstrModel
End Property

Public Property Let Model(ByVal value As String)
    mstrModel = value
End Property

Public Property Set messages(ByVal value As clsOpenAIMessages)
    Set mobjMessages = value
End Property

Public Property Get MaxTokens() As Long
    MaxTokens = mlngMaxTokens
End Property

Public Property Let MaxTokens(ByVal value As Long)
    mlngMaxTokens = value
End Property

Public Property Let TopP(ByVal value As Double)
    mdblTopP = value
End Property

Public Property Let Temperature(ByVal value As Double)
    mdblTemperature = value
End Property

Public Property Get Temperature() As Double
    Temperature = mdblTemperature
End Property

Public Property Let FrequencyPenalty(ByVal value As Double)
    mdblFrequencyPenalty = value
End Property

Public Property Let PresencePenalty(ByVal value As Double)
    mdlPresencePenalty = value
End Property

Public Property Get ImageHeight() As Long
    ImageHeight = mlngImageHeight
End Property

Public Property Let ImageHeight(ByVal value As Long)
    mlngImageHeight = value
End Property

Public Property Get ImageWidth() As Long
    ImageWidth = mlngImageHeight
End Property

Public Property Let ImageWidth(ByVal value As Long)
    mlngImageWidth = value
End Property

Public Property Let TimeoutResolve(ByVal value As Long)
    mlngTimeoutResolve = value
End Property

Public Property Get TimeoutResolve() As Long
    TimeoutResolve = mlngTimeoutResolve
End Property

Public Property Let TimeoutConnect(ByVal value As Long)
    mlngTimeoutConnect = value
End Property

Public Property Get TimeoutConnect() As Long
    TimeoutConnect = mlngTimeoutConnect
End Property

Public Property Let TimeoutSend(ByVal value As Long)
    mlngTimeoutSend = value
End Property

Public Property Get TimeoutSend() As Long
    TimeoutSend = mlngTimeoutSend
End Property

Public Property Let TimeoutReceive(ByVal value As Long)
    mlngTimeoutReceive = value
End Property

Public Property Get TimeoutReceive() As Long
    TimeoutReceive = mlngTimeoutReceive
End Property


Public Function GetChatSendToAPIJsonString() As String
    GetChatSendToAPIJsonString = "{""model"": """ & mstrModel & """, " & mobjMessages.GetAllMessages & ", ""max_tokens"": " & mlngMaxTokens & ", ""top_p"": " & mdblTopP & ", ""temperature"": " & mdblTemperature & ", ""frequency_penalty"": " & mdblFrequencyPenalty & ", ""presence_penalty"": " & mdlPresencePenalty & "}"
End Function


Public Function GetTextCompletionSendToAPIJsonString() As String
    GetTextCompletionSendToAPIJsonString = "{""model"": """ & mstrModel & """, ""prompt"": """ & mstrPrompt & """, ""max_tokens"": " & mlngMaxTokens & ", ""temperature"": " & mdblTemperature & "}"
End Function


Public Function GetDalleImageSendToAPIJsonString() As String
    Dim strImageSize As String
    strImageSize = Me.GetImageDimensionLabel()
    GetDalleImageSendToAPIJsonString = "{""prompt"": """ & mstrPrompt & """, ""size"": " & strImageSize & "}"
End Function


Public Function GetImageDimensionLabel() As String
    GetImageDimensionLabel = Chr(34) & CStr(mlngImageWidth) & "x" & CStr(mlngImageHeight) & Chr(34)
End Function
