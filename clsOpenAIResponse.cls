VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenAIResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Project: OpenAI VBA Framework
' Class:  clsOpenAIResponse
' Description: Handles and formats the results json from the OpenAI API
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

Private mstrId As String
Private mstrObject As String
Private mstrCreated As Long
Private mstrModel As String
Private mintPromptTokens As Integer
Private mintCompletionTokens As Integer
Private mintTotalTokens As Integer
Private mstrMessageRole As String
Private mstrMessageContent As String
Private mstrFinishReason As String
Private mintIndex As Integer
Private mstrText As String
Private mstrLogprobs As String
Private mstrSavedFile As String
Private mstrJson As String

Private Function IOpenAINameProvider_GetClassName() As String
    IOpenAINameProvider_GetClassName = "clsOpenAIResponse"
End Function


Private Function IOpenAINameProvider_ToString() As String
    
    Dim strConcatenatedString As String

    strConcatenatedString = mstrId & ", " & mstrObject & ", " & CStr(mstrCreated) & ", " & mstrModel & ", " & CStr(mintPromptTokens) & ", " & CStr(mintCompletionTokens) & ", " & CStr(mintTotalTokens) & ", " & mstrMessageRole & ", " & mstrMessageContent & ", " & mstrFinishReason & ", " & CStr(mintIndex)
    
    IOpenAINameProvider_ToString = "values: " & strConcatenatedString

End Function


Public Property Get Json() As String
    Json = mstrJson
End Property

Public Property Get Id() As String
    Id = mstrId
End Property

Public Property Get Object() As String
    Object = mstrObject
End Property

Public Property Get Created() As Long
    Created = mstrCreated
End Property

Public Property Get Model() As String
    Model = mstrModel
End Property

Public Property Get PromptTokens() As Integer
    PromptTokens = mintPromptTokens
End Property

Public Property Get CompletionTokens() As Integer
    CompletionTokens = mintCompletionTokens
End Property

Public Property Get TotalTokens() As Integer
    TotalTokens = mintTotalTokens
End Property

Public Property Get MessageRole() As String
    MessageRole = mstrMessageRole
End Property

Public Property Get MessageContent() As String
    MessageContent = mstrMessageContent
End Property

Public Property Get TextContent() As String
    TextContent = mstrText
End Property

Public Property Get LogProbs() As String
    LogProbs = mstrLogprobs
End Property

Public Property Get FinishReason() As String
    FinishReason = mstrFinishReason
End Property

Public Property Get Index() As Integer
    Index = mintIndex
End Property

Public Property Get SavedLocalFile() As String
    SavedLocalFile = mstrSavedFile
End Property

Public Property Let SavedLocalFile(ByVal value As String)
    mstrSavedFile = value
End Property

Public Function IsExistSavedLocalFile() As Boolean
    IsExistSavedLocalFile = IIf(Len(Dir(mstrSavedFile)) > 0, True, False)
End Function


Private Sub Class_Initialize()
' Initialize the variables
    mstrJson = Empty
    mstrId = Empty
    mstrObject = Empty
    mstrCreated = 0
    mstrModel = Empty
    mintPromptTokens = 0
    mintCompletionTokens = 0
    mintTotalTokens = 0
    mstrMessageRole = Empty
    mstrMessageContent = Empty
    mstrFinishReason = Empty
    mintIndex = 0
    mstrText = Empty
    mstrLogprobs = Empty
    mstrSavedFile = Empty
End Sub


Public Sub ParseChatJSON(ByVal strJson As String)
    ' Purpose: This method is for parsing OpenAI's strJson response from its Chat Endpoint

    'Allow for the full response to be accessed by downstream clients using the class
    mstrJson = strJson

    ' Extract each property from the JSON string
    mstrId = ExtractJsonValue(strJson, """id"": """, """")
    mstrObject = ExtractJsonValue(strJson, """object"": """, """")
    mstrCreated = CLng(ExtractJsonValue(strJson, """created"": ", ","))
    mstrModel = ExtractJsonValue(strJson, """model"": """, """")
    mintPromptTokens = CInt(ExtractJsonValue(strJson, """prompt_tokens"": ", ","))
    mintCompletionTokens = CInt(ExtractJsonValue(strJson, """completion_tokens"": ", ","))
    mintTotalTokens = CInt(ExtractJsonValue(strJson, """total_tokens"": ", "}"))
    
    ' Extract the nested message information
    Dim messageJson As String
    messageJson = ExtractJsonValue(strJson, """message"": {", "}")
    mstrMessageRole = ExtractJsonValue(messageJson, """role"": """, """")
    mstrMessageContent = ExtractJsonValue(messageJson, """content"": """, """")
    mstrMessageContent = Replace(mstrMessageContent, "\""", """") ' Replace escaped quotes with actual quotes
    mstrFinishReason = ExtractJsonValue(strJson, """finish_reason"": """, """")
    mintIndex = CInt(ExtractJsonValue(strJson, """index"": ", ","))

End Sub


Private Function ExtractJsonValue(ByVal Json As String, ByVal key As String, ByVal delimiter As String) As String
    ' Find the start position of the key
    Dim startPos As Integer
    startPos = InStr(Json, key)
    If startPos = 0 Then Exit Function

    ' Adjust start position to start of the value
    startPos = startPos + Len(key)

    ' Find the end position of the value
    Dim endPos As Integer
    endPos = InStr(startPos, Json, delimiter)
    If endPos = 0 Then Exit Function

    ' Extract the value
    ExtractJsonValue = Mid(Json, startPos, endPos - startPos)
End Function


Public Function GetFileNameFromImageURL(ByVal strImageUrl As String)

    GetFileNameFromImageURL = Empty
    
    If strImageUrl <> Empty Then

        Dim intFileNameStartPos As Integer
        intFileNameStartPos = InStr(1, strImageUrl, "img-")
        
        Dim intFileNameEndPos As Integer
        intFileNameEndPos = InStr(1, strImageUrl, "png") + 3
        
        Dim intFileNameLength As Integer
        intFileNameLength = intFileNameEndPos - intFileNameStartPos
    
        Dim strFileName As String
        strFileName = Mid(strImageUrl, intFileNameStartPos, intFileNameLength)
        
        GetFileNameFromImageURL = strFileName

    End If
    
End Function


Public Function GetImageURLFromImageCreationJSON(ByVal strResponseJson As String)
'Purpose: This method is for parsing OpenAI's json from it's text completion end point

    GetImageURLFromImageCreationJSON = Empty
    
    If strResponseJson <> Empty Then
    
        Dim intStartPos As Integer
        intStartPos = InStr(1, strResponseJson, Chr(34) & "url" & Chr(34)) + 8
    
        Dim intEndPos As Integer
        intEndPos = InStr(1, strResponseJson, "}") - 6
        
        Dim intLength As Integer
        intLength = intEndPos - intStartPos
        
        Dim strImageUrl As String
        strImageUrl = Mid(strResponseJson, intStartPos, intLength)
        
        GetImageURLFromImageCreationJSON = strImageUrl
    
    End If
    
End Function

