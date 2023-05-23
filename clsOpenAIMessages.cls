VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenAIMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Project: OpenAI VBA Framework
' Class:  clsOpenAIMessages
' Description: Holds the message objects (clsOpenAIMessage) to send to Chat API
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

Private mcollMessages As Collection


Private Function IOpenAINameProvider_GetClassName() As String
    IOpenAINameProvider_GetClassName = "clsOpenAIMessages"
End Function

Private Function IOpenAINameProvider_ToString() As String
    IOpenAINameProvider_ToString = GetAllMessages
End Function


Private Sub Class_Initialize()
    Set mcollMessages = New Collection
End Sub


Private Sub Class_Terminate()
    Set mcollMessages = Nothing
End Sub


Public Sub AddMessage(ByVal objMsg As clsOpenAIMessage)
    mcollMessages.Add objMsg
End Sub


Public Sub ClearMessages()
    If Not mcollMessages Is Nothing Then
        mcollMessages.Clear
    End If
End Sub


Public Function GetAllMessages() As String
'Purpose: This builds a string of all messages that will be sent to OpenAI API

    Dim objMsg As clsOpenAIMessage
    Dim strOutput As String
    
    If Me.GetTotal = 0 Then
        GetAllMessages = ""
        Exit Function
    End If
    
    strOutput = """messages"": ["
    
    For Each objMsg In mcollMessages
        strOutput = strOutput & objMsg.GetMessageContentString & ", "
    Next objMsg
    
    ' Remove the trailing comma and space
    If mcollMessages.Count > 0 Then
        strOutput = Left(strOutput, Len(strOutput) - 2)
    End If
    
    strOutput = strOutput & "]"
    
    GetAllMessages = strOutput
    
    Set objMsg = Nothing
    
End Function


Private Function GetMessageObject(ByVal strKeyName As String, ByVal strKey As String, ByVal strValueName As String, ByVal strValue As String) As clsOpenAIMessage
'Purpose: Helper routine to construct known message object types
    
    Dim objMsg As clsOpenAIMessage
    Set objMsg = New clsOpenAIMessage
    
    objMsg.Add strKeyName, strKey, strValueName, strValue
    
    Set GetMessageObject = objMsg
    Set objMsg = Nothing

End Function


Public Sub AddUserMessage(ByVal strContent As String)
    Call Me.AddMessage(GetMessageObject("role", "user", "content", strContent))
End Sub


Public Sub AddSystemMessage(ByVal strContent As String)
    Call Me.AddMessage(GetMessageObject("role", "system", "content", strContent))
End Sub


Public Sub AddAssistantMessage(ByVal strContent As String)
    Call Me.AddMessage(GetMessageObject("role", "assistant", "content", strContent))
End Sub


Public Sub AddCustomMessage(ByVal strKeyName As String, ByVal strKeyValue As String, ByVal strValueName As String, ByVal strValue As String)
    Call Me.AddMessage(GetMessageObject(strKeyName, strKeyValue, strValueName, strValue))
End Sub


Public Function GetTotal() As Integer
    If Not mcollMessages Is Nothing Then
        GetTotal = mcollMessages.Count
    Else
        GetTotal = 0
    End If
End Function


Public Function IsPopulated() As Boolean
    IsPopulated = IIf(Me.GetTotal > 0, True, False)
End Function
