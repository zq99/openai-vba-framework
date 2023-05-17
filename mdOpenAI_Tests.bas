Attribute VB_Name = "mdOpenAI_TESTS"
'-----------------------------------------------------------------------------
' Project: OpenAI VBA Framework
' Module:  mdOpenAI_Tests
' Description: Tests the framework is retrieving data correctly from OpenAI
'
' Author: Zaid Qureshi
' GitHub: https://github.com/zq99
'
' Modules in the Framework:
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

' This work is licensed under the MIT License. The full license text
' can be found in the LICENSE file in the root of this repository.
'
' Copyright (c) 2023 Zaid Qureshi
'-----------------------------------------------------------------------------

Option Explicit

'******************************************************
' GET YOUR API KEY: https://openai.com/api/
Public Const API_KEY As String = "<API_KEY>"
'******************************************************


Public Sub TestOpenAI()
'Purpose: This tests all endpoints are being queried correctly and returning data

    Dim oOpenAI As clsOpenAI
    Dim oMessages As New clsOpenAIMessages
    Dim oResponse As clsOpenAIResponse
    
    Set oOpenAI = New clsOpenAI
    
    'All output to sent to immediate window
    oOpenAI.IsLogOutputRequired True
    oOpenAI.API_KEY = API_KEY
    oOpenAI.Temperature = 0
    
    oOpenAI.Log "(1) Simple chat test"
    
    oMessages.AddSystemMessage "Every answer should only contain alphanumeric characters, and every letter should be capitalized"
    oMessages.AddUserMessage "What is the capital of France?"

    Set oResponse = oOpenAI.ChatCompletion(oMessages)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.MessageContent) > 0
    Debug.Assert oResponse.MessageContent = "PARIS"
    Debug.Assert oResponse.MessageRole = "assistant"
    
    oOpenAI.Log oMessages.GetAllMessages
    oOpenAI.Log oResponse.MessageContent
    oOpenAI.Log oResponse.MessageRole
    
    oOpenAI.Log "(2) Simple chat test with temperature change"

    oMessages.AddUserMessage "write a string of digits in order up to 9"

    Set oResponse = oOpenAI.ChatCompletion(oMessages)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.MessageContent) > 0
    Debug.Assert oResponse.MessageContent = "123456789"
    Debug.Assert oResponse.MessageRole = "assistant"
    
    '(3) Text completion test
    
    Dim strMsg As String
    
    'reset to default
    oOpenAI.ClearSettings
    
    strMsg = "Write a Haiku about a dinosaur that loves to code in VBA"
    Set oResponse = oOpenAI.TextCompletion(strMsg)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.TextContent) > 0
    oOpenAI.Log (oResponse.TextContent)
    
    
    Set oResponse = Nothing
    Set oOpenAI = Nothing
    Set oMessages = Nothing

End Sub
