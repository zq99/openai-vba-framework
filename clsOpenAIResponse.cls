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
' Copyright (c) 2023 Zaid Qureshi
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


Private Function IOpenAINameProvider_GetClassName() As String
    IOpenAINameProvider_GetClassName = "clsOpenAIResponse"
End Function

Private Function IOpenAINameProvider_ToString() As String
    Dim strConcatenatedString As String

    strConcatenatedString = mstrId & ", " & mstrObject & ", " & CStr(mstrCreated) & ", " & mstrModel & ", " & CStr(mintPromptTokens) & ", " & CStr(mintCompletionTokens) & ", " & CStr(mintTotalTokens) & ", " & mstrMessageRole & ", " & mstrMessageContent & ", " & mstrFinishReason & ", " & CStr(mintIndex)
    
    IOpenAINameProvider_ToString = "values: " & strConcatenatedString
End Function

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

Public Sub ParseChatJSON(ByVal json As String)
'Purpose: This method is for parsing OpenAI's json response from it's Chat End point

    Dim intStartPos As Integer
    Dim intEndPos As Integer
    Dim strTemp As String
    
    ' Extract "id"
    If InStr(1, json, """id"":""") > 0 Then
        intStartPos = InStr(1, json, """id"":""") + Len("""id"":""")
        intEndPos = InStr(intStartPos, json, """")
        mstrId = Trim(Mid(json, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "object"
    If InStr(1, json, """object"":""") > 0 Then
        intStartPos = InStr(1, json, """object"":""") + Len("""object"":""")
        intEndPos = InStr(intStartPos, json, """")
        mstrObject = Trim(Mid(json, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "created"
    If InStr(1, json, """created"":") > 0 Then
        intStartPos = InStr(1, json, """created"":") + Len("""created"":")
        intEndPos = InStr(intStartPos, json, ",")
        mstrCreated = CLng(Trim(Mid(json, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "model"
    If InStr(1, json, """model"":""") > 0 Then
        intStartPos = InStr(1, json, """model"":""") + Len("""model"":""")
        intEndPos = InStr(intStartPos, json, """")
        mstrModel = Trim(Mid(json, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "prompt_tokens"
    If InStr(1, json, """prompt_tokens"":") > 0 Then
        intStartPos = InStr(1, json, """prompt_tokens"":") + Len("""prompt_tokens"":")
        intEndPos = InStr(intStartPos, json, ",")
        mintPromptTokens = CInt(Trim(Mid(json, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "completion_tokens"
    If InStr(1, json, """completion_tokens"":") > 0 Then
        intStartPos = InStr(1, json, """completion_tokens"":") + Len("""completion_tokens"":")
        intEndPos = InStr(intStartPos, json, ",")
        mintCompletionTokens = CInt(Trim(Mid(json, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "total_tokens"
    If InStr(1, json, """total_tokens"":") > 0 Then
        intStartPos = InStr(1, json, """total_tokens"":") + Len("""total_tokens"":")
        intEndPos = InStr(intStartPos, json, "}")
        mintTotalTokens = CInt(Trim(Mid(json, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "message_role"
    If InStr(1, json, """role"":""") > 0 Then
        intStartPos = InStr(1, json, """role"":""") + Len("""role"":""")
        intEndPos = InStr(intStartPos, json, """")
        mstrMessageRole = Trim(Mid(json, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "message_content"
    If InStr(1, json, """content"":""") > 0 Then
        intStartPos = InStr(1, json, """content"":""") + Len("""content"":""")
        intEndPos = InStr(intStartPos, json, """},") ' end position is now before "}," sequence
        strTemp = Mid(json, intStartPos, intEndPos - intStartPos)
        strTemp = Replace(strTemp, "\""", """") ' Replace escaped quotes with actual quotes
        mstrMessageContent = Trim(strTemp)
    End If

    
    ' Extract "finish_reason"
    If InStr(1, json, """finish_reason"":""") > 0 Then
        intStartPos = InStr(1, json, """finish_reason"":""") + Len("""finish_reason"":""")
        intEndPos = InStr(intStartPos, json, """")
        mstrFinishReason = Trim(Mid(json, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "index"
    If InStr(1, json, """index"":") > 0 Then
        intStartPos = InStr(1, json, """index"":") + Len("""index"":")
        intEndPos = InStr(intStartPos, json, "}")
        mintIndex = CInt(Trim(Mid(json, intStartPos, intEndPos - intStartPos)))
    End If
End Sub



Public Sub ParseTextCompletionJSON(ByVal json As String)
'Purpose: This method is for parsing OpenAI's json from it's text completion end point

    Dim intStartPos As Integer
    Dim intEndPos As Integer
    Dim strTemp As String
    
    ' Extract "id"
    If InStr(1, json, """id"":""") > 0 Then
        intStartPos = InStr(1, json, """id"":""") + Len("""id"":""")
        intEndPos = InStr(intStartPos, json, """")
        mstrId = Trim(Mid(json, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "object"
    If InStr(1, json, """object"":""") > 0 Then
        intStartPos = InStr(1, json, """object"":""") + Len("""object"":""")
        intEndPos = InStr(intStartPos, json, """")
        mstrObject = Trim(Mid(json, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "created"
    If InStr(1, json, """created"":") > 0 Then
        intStartPos = InStr(1, json, """created"":") + Len("""created"":")
        intEndPos = InStr(intStartPos, json, ",")
        mstrCreated = CLng(Trim(Mid(json, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "model"
    If InStr(1, json, """model"":""") > 0 Then
        intStartPos = InStr(1, json, """model"":""") + Len("""model"":""")
        intEndPos = InStr(intStartPos, json, """")
        mstrModel = Trim(Mid(json, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "prompt_tokens"
    If InStr(1, json, """prompt_tokens"":") > 0 Then
        intStartPos = InStr(1, json, """prompt_tokens"":") + Len("""prompt_tokens"":")
        intEndPos = InStr(intStartPos, json, ",")
        mintPromptTokens = CInt(Trim(Mid(json, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "completion_tokens"
    If InStr(1, json, """completion_tokens"":") > 0 Then
        intStartPos = InStr(1, json, """completion_tokens"":") + Len("""completion_tokens"":")
        intEndPos = InStr(intStartPos, json, ",")
        mintCompletionTokens = CInt(Trim(Mid(json, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "total_tokens"
    If InStr(1, json, """total_tokens"":") > 0 Then
        intStartPos = InStr(1, json, """total_tokens"":") + Len("""total_tokens"":")
        intEndPos = InStr(intStartPos, json, "}")
        mintTotalTokens = CInt(Trim(Mid(json, intStartPos, intEndPos - intStartPos)))
    End If
    
        ' Extract "text"
    If InStr(1, json, """text"":""") > 0 Then
        intStartPos = InStr(1, json, """text"":""") + Len("""text"":""")
        intEndPos = InStr(intStartPos, json, """,""") ' end position is now before the sequence ","
        strTemp = Mid(json, intStartPos, intEndPos - intStartPos)
        strTemp = Replace(strTemp, "\""", """") ' Replace escaped quotes with actual quotes
        mstrText = Trim(strTemp)
    End If
    
    ' Extract "logprobs"
    If InStr(1, json, """logprobs"":") > 0 Then
        intStartPos = InStr(1, json, """logprobs"":") + Len("""logprobs"":")
        intEndPos = InStr(intStartPos, json, ",") ' end position is now before the sequence ","
        strTemp = Mid(json, intStartPos, intEndPos - intStartPos)
        strTemp = Replace(strTemp, "\""", """") ' Replace escaped quotes with actual quotes
        mstrLogprobs = Trim(strTemp)
    End If

    
    ' Extract "finish_reason"
    If InStr(1, json, """finish_reason"":""") > 0 Then
        intStartPos = InStr(1, json, """finish_reason"":""") + Len("""finish_reason"":""")
        intEndPos = InStr(intStartPos, json, """")
        mstrFinishReason = Trim(Mid(json, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "index"
    If InStr(1, json, """index"":") > 0 Then
        intStartPos = InStr(1, json, """index"":") + Len("""index"":")
        intEndPos = InStr(intStartPos, json, ",")
        mintIndex = CInt(Trim(Mid(json, intStartPos, intEndPos - intStartPos)))
    End If
End Sub

