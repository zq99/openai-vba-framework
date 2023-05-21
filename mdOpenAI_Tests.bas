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

#If VBA7 Then
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'******************************************************
' GET YOUR API KEY: https://openai.com/api/
Public Const API_KEY As String = "<API_KEY>"
'******************************************************


Public Sub RunAllTests()
'********************************************************************************
'Purpose: This tests all endpoints are being queried correctly and returning data
'********************************************************************************

    Dim arrMSXMLTypes(1 To 3) As String
    Dim oOpenAI As New clsOpenAI
    
    oOpenAI.IsLogOutputRequired True
    oOpenAI.API_KEY = API_KEY

    ' Assign all posssible MSXML types
    arrMSXMLTypes(1) = Empty
    arrMSXMLTypes(2) = oOpenAI.MSXML_XML_VALUE
    arrMSXMLTypes(3) = oOpenAI.MSXML_SERVER_XML_VALUE

    ' Declare a variable for the loop index
    Dim i As Integer

    ' Loop through each item in the array
    For i = LBound(arrMSXMLTypes) To UBound(arrMSXMLTypes)
        DoEvents
        oOpenAI.Log arrMSXMLTypes(i)
        Call TestOpenAI(oOpenAI, arrMSXMLTypes(i))
        Sleep 1000
    Next i

    Set oOpenAI = Nothing

End Sub


Private Sub TestOpenAI(ByVal oOpenAI As clsOpenAI, Optional ByVal strRequestXMLType As String)

    Dim oMessages As New clsOpenAIMessages
    Dim oResponse As clsOpenAIResponse
        
    If strRequestXMLType <> Empty Then
        oOpenAI.MSXMLType = oOpenAI.MSXML_SERVER_XML_VALUE
    End If
    
    'All output to sent to immediate window
    oOpenAI.Temperature = 0
    
    '*********************************************
    '(1) Simple chat test
    '*********************************************
    
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
    
    '*********************************************
    '(2) Simple chat test with temperature change
    '*********************************************

    oMessages.AddUserMessage "write a string of digits in order up to 9"
    oOpenAI.Temperature = 0.9
    Set oResponse = oOpenAI.ChatCompletion(oMessages)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.MessageContent) > 0
    Debug.Assert oResponse.MessageContent = "123456789"
    Debug.Assert oResponse.MessageRole = "assistant"
    
    '*********************************************
    '(3) Change timeouts
    '*********************************************

    oMessages.AddUserMessage "write a string of digits in order up to 9"
    oOpenAI.SetTimeOutDefaults 5000, 5000, 5000, 5000
    Set oResponse = oOpenAI.ChatCompletion(oMessages)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.MessageContent) > 0
    Debug.Assert oResponse.MessageContent = "123456789"
    Debug.Assert oResponse.MessageRole = "assistant"
    
    '*********************************************
    '(4) Text completion test
    '*********************************************
    
    Dim strMsg As String
    
    'reset to default
    oOpenAI.ClearSettings
    
    strMsg = "Write a Haiku about a dinosaur that loves to code in VBA"
    Set oResponse = oOpenAI.TextCompletion(strMsg)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.TextContent) > 0
    oOpenAI.Log (oResponse.TextContent)
    
    '*********************************************
    '(5) Image creation from prompt test
    '*********************************************
    
    oOpenAI.ClearSettings
    Set oResponse = oOpenAI.CreateImageFromText("A cat playing a banjo on a surfboard", 256, 256)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.SavedLocalFile) > 0
    Debug.Assert Len(Dir(oResponse.SavedLocalFile)) > 0
    
    Set oResponse = Nothing
    Set oMessages = Nothing

End Sub
