Attribute VB_Name = "mdOpenAI_Examples"
'-----------------------------------------------------------------------------
' Project: OpenAI VBA Framework
' Module:  mdOpenAI_Examples
' Description: Tests all the framework is retrieving data correctly
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
'
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

Public Sub TestSimpleOpenAI()

    Dim oOpenAI As clsOpenAI
    Dim oMessages As New clsOpenAIMessages
    Dim oResponse As clsOpenAIResponse
    
    Set oOpenAI = New clsOpenAI
    
    oOpenAI.API_KEY = API_KEY
    
    oMessages.AddSystemMessage "You must always have a sarcastic response to every question you are asked and never give any practical advice."
    oMessages.AddUserMessage "Hey man do you know how to get to Carnegie Hall?"

    Set oResponse = oOpenAI.ChatCompletion(oMessages)
    If Not oResponse Is Nothing Then
        Debug.Print (oResponse.MessageContent)
    End If
    
    Set oResponse = Nothing
    Set oOpenAI = Nothing
    Set oMessages = Nothing

End Sub


Public Sub TestChatOpenAI()

    Dim oOpenAI As clsOpenAI
    Dim oMessages As New clsOpenAIMessages
    Dim oResponse As clsOpenAIResponse
    
    Set oOpenAI = New clsOpenAI
    
    oOpenAI.Model = "gpt-4"
    
    oOpenAI.API_KEY = API_KEY
    
    oMessages.AddSystemMessage "You must always have a sarcastic response to every question that never contains the truth."
    oMessages.AddUserMessage "Hey man do you know how to get to Carnegie Hall?"

    If oMessages.IsPopulated Then
        Set oResponse = oOpenAI.ChatCompletion(oMessages)
        If Not oResponse Is Nothing Then
            Debug.Print (oResponse.Id)
            Debug.Print (oResponse.Object)
            Debug.Print (oResponse.Created)
            Debug.Print (oResponse.Model)
            Debug.Print (oResponse.FinishReason)
            Debug.Print (oResponse.CompletionTokens)
            Debug.Print (oResponse.MessageRole)
            Debug.Print (oResponse.MessageContent)
            Debug.Print (oResponse.PromptTokens)
            Debug.Print (oResponse.TotalTokens)
            Debug.Print (oResponse.Index)
        End If
    End If
    
    Set oResponse = Nothing
    Set oOpenAI = Nothing
    Set oMessages = Nothing

End Sub


Public Sub TestTextCompletionOpenAI()

    Dim oOpenAI As clsOpenAI
    Dim oResponse As clsOpenAIResponse
    Dim sMsg As String
    
    Set oOpenAI = New clsOpenAI
    
    oOpenAI.IsLogOutputRequired True
    
    oOpenAI.API_KEY = API_KEY

    sMsg = "Write a joke about a dinosaur attempting to learn how to salsa dance"
    Set oResponse = oOpenAI.TextCompletion(sMsg)
    
    If Not oResponse Is Nothing Then
        Debug.Print (oResponse.Id)
        Debug.Print (oResponse.Object)
        Debug.Print (oResponse.Created)
        Debug.Print (oResponse.Model)
        Debug.Print (oResponse.FinishReason)
        Debug.Print (oResponse.TextContent)
        Debug.Print (oResponse.LogProbs)
        Debug.Print (oResponse.CompletionTokens)
        Debug.Print (oResponse.PromptTokens)
        Debug.Print (oResponse.TotalTokens)
        Debug.Print (oResponse.Index)
    End If
    
    Set oResponse = Nothing
    Set oOpenAI = Nothing

End Sub


Public Sub TestTextCompletionSimpleOpenAI()

    Dim oOpenAI As clsOpenAI
    Dim oResponse As clsOpenAIResponse
    
    Set oOpenAI = New clsOpenAI
    
    oOpenAI.API_KEY = API_KEY

    Set oResponse = oOpenAI.TextCompletion("Write a Haiku about a Dinosaur that loves to code!")
    
    If Not oResponse Is Nothing Then
        Debug.Print (oResponse.TextContent)
    End If
    
    Set oResponse = Nothing
    Set oOpenAI = Nothing

End Sub



Public Function GETTEXTFROMOPENAI(prompt As String, apiKey As String, Optional ByVal Model) As String
    Dim oOpenAI As clsOpenAI
    Dim oResponse As clsOpenAIResponse

    Set oOpenAI = New clsOpenAI

    ' Set the API key directly from the function argument
    oOpenAI.API_KEY = apiKey
    
    If Not IsEmpty(Model) Then
        oOpenAI.Model = Model
    End If

    ' Make the API request and get the response
    Set oResponse = oOpenAI.TextCompletion(prompt)

    ' Return the choice from the response, or an empty string if there was no response
    If Not oResponse Is Nothing Then
        GETTEXTFROMOPENAI = oResponse.TextContent
    Else
        GETTEXTFROMOPENAI = ""
    End If
End Function

