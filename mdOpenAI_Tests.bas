Attribute VB_Name = "mdOpenAI_TESTS"
'-----------------------------------------------------------------------------
' Project: OpenAI VBA Framework
' Module:  mdOpenAI_Tests
' Description: Tests the framework is retrieving data correctly from OpenAI
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

' This work is licensed under the MIT License. The full license text
' can be found in the LICENSE file in the root of this repository.
'
'-----------------------------------------------------------------------------

Option Explicit

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
    
    Debug.Assert oOpenAI.API_KEY = API_KEY
    Debug.Assert oOpenAI.CallsToAPICount = 0

    ' Assign all posssible MSXML types
    arrMSXMLTypes(1) = Empty
    arrMSXMLTypes(2) = oOpenAI.MSXML_XML_VALUE
    arrMSXMLTypes(3) = oOpenAI.MSXML_SERVER_XML_VALUE

    Dim i As Integer

    ' Loop through each item in the array
    For i = LBound(arrMSXMLTypes) To UBound(arrMSXMLTypes)
        DoEvents
        oOpenAI.Log arrMSXMLTypes(i)
        Call TestOpenAI(oOpenAI, arrMSXMLTypes(i))
        oOpenAI.Pause
    Next i
    
    Debug.Assert oOpenAI.CallsToAPICount > 0
    Set oOpenAI = Nothing

End Sub


Private Sub TestOpenAI(ByVal oOpenAI As clsOpenAI, Optional ByVal strRequestXMLType As String)

    Dim oMessages As New clsOpenAIMessages
    Dim oResponse As clsOpenAIResponse
        
    If strRequestXMLType <> Empty Then
        oOpenAI.MSXMLType = strRequestXMLType
        Debug.Assert oOpenAI.MSXMLType = strRequestXMLType
    End If
    
    'Test temperature can be changed
    oOpenAI.Temperature = 0.9
    Debug.Assert oOpenAI.Temperature = 0.9
    
    'Set least amount of variation for testing
    oOpenAI.Temperature = 0
    Debug.Assert oOpenAI.Temperature = 0
    
    '*********************************************
    '(1) Simple chat test
    '*********************************************
    
    'Test with different models
    Dim arrModels(1 To 2) As String
    Dim i As Integer
    
    arrModels(1) = Empty
    arrModels(2) = "gpt-4"
    
    For i = LBound(arrModels) To UBound(arrModels)
    
        DoEvents
    
        If arrModels(i) <> Empty Then
            oOpenAI.Model = arrModels(i)
            Debug.Assert oOpenAI.Model = arrModels(i)
        End If
        
        oOpenAI.Log "Testing with model: " & IIf(oOpenAI.Model = Empty, "[None]", oOpenAI.Model)
    
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
        
        oOpenAI.Pause 5000
        
    Next i
    
    '*********************************************
    '(2) Simple chat test with temperature change
    '*********************************************

    oMessages.AddUserMessage "write a string of digits in order up to 9 starting with 1 and ending with 9"
    Set oResponse = oOpenAI.ChatCompletion(oMessages)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.MessageContent) > 0
    Debug.Assert oResponse.MessageContent = "123456789"
    Debug.Assert oResponse.MessageRole = "assistant"
    
    oOpenAI.Pause 5000
    
    '*********************************************
    '(3) Change timeouts
    '*********************************************

    oMessages.AddUserMessage "write a string of digits in order up to 9 starting with 1 and ending with 9"
    oOpenAI.SetTimeOutDefaults 5000, 5000, 5000, 5000
    Set oResponse = oOpenAI.ChatCompletion(oMessages)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.MessageContent) > 0
    Debug.Assert oResponse.MessageContent = "123456789"
    Debug.Assert oResponse.MessageRole = "assistant"
    
    oOpenAI.Pause 5000
    
    '*********************************************
    '(4) Text completion test
    '*********************************************
    
    Dim strMsg As String
    
    'reset to default
    oOpenAI.ClearSettings
    Debug.Assert oOpenAI.CallsToAPICount = 0
    
    strMsg = "Write a Haiku about a dinosaur that loves to code in VBA"
    Set oResponse = oOpenAI.TextCompletion(strMsg)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.TextContent) > 0
    oOpenAI.Log (oResponse.TextContent)
    
    oOpenAI.Pause 5000
    
    '*********************************************
    '(5) Image creation from prompt test
    '*********************************************
    
    oOpenAI.ClearSettings
    
    strMsg = "A cat playing a banjo on a surfboard"
    Set oResponse = oOpenAI.CreateImageFromText(strMsg, 256, 256)
    
    Debug.Assert Not oResponse Is Nothing
    Debug.Assert Len(oResponse.SavedLocalFile) > 0
    Debug.Assert Len(Dir(oResponse.SavedLocalFile)) > 0
    Debug.Assert oResponse.IsExistSavedLocalFile = True
    
    oOpenAI.Log ("Prompt=" & strMsg)
    oOpenAI.Log ("Image saved to: " & oResponse.SavedLocalFile)
    oOpenAI.Pause 5000
    
    Set oResponse = Nothing
    Set oMessages = Nothing

End Sub
