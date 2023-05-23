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

Public Property Get SavedLocalFile() As String
    SavedLocalFile = mstrSavedFile
End Property

Public Property Let SavedLocalFile(ByVal value As String)
    mstrSavedFile = value
End Property


Private Sub Class_Initialize()
' Initialize the variables
    mstrId = ""
    mstrObject = ""
    mstrCreated = 0
    mstrModel = ""
    mintPromptTokens = 0
    mintCompletionTokens = 0
    mintTotalTokens = 0
    mstrMessageRole = ""
    mstrMessageContent = ""
    mstrFinishReason = ""
    mintIndex = 0
    mstrText = ""
    mstrLogprobs = ""
    mstrSavedFile = ""
End Sub


Public Sub ParseChatJSON(ByVal strJson As String)
'Purpose: This method is for parsing OpenAI's strJson response from it's Chat End point

    Dim intStartPos As Integer
    Dim intEndPos As Integer
    Dim strTemp As String
    
    ' Extract "id"
    If InStr(1, strJson, """id"":""") > 0 Then
        intStartPos = InStr(1, strJson, """id"":""") + Len("""id"":""")
        intEndPos = InStr(intStartPos, strJson, """")
        mstrId = Trim(Mid(strJson, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "object"
    If InStr(1, strJson, """object"":""") > 0 Then
        intStartPos = InStr(1, strJson, """object"":""") + Len("""object"":""")
        intEndPos = InStr(intStartPos, strJson, """")
        mstrObject = Trim(Mid(strJson, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "created"
    If InStr(1, strJson, """created"":") > 0 Then
        intStartPos = InStr(1, strJson, """created"":") + Len("""created"":")
        intEndPos = InStr(intStartPos, strJson, ",")
        mstrCreated = CLng(Trim(Mid(strJson, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "model"
    If InStr(1, strJson, """model"":""") > 0 Then
        intStartPos = InStr(1, strJson, """model"":""") + Len("""model"":""")
        intEndPos = InStr(intStartPos, strJson, """")
        mstrModel = Trim(Mid(strJson, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "prompt_tokens"
    If InStr(1, strJson, """prompt_tokens"":") > 0 Then
        intStartPos = InStr(1, strJson, """prompt_tokens"":") + Len("""prompt_tokens"":")
        intEndPos = InStr(intStartPos, strJson, ",")
        mintPromptTokens = CInt(Trim(Mid(strJson, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "completion_tokens"
    If InStr(1, strJson, """completion_tokens"":") > 0 Then
        intStartPos = InStr(1, strJson, """completion_tokens"":") + Len("""completion_tokens"":")
        intEndPos = InStr(intStartPos, strJson, ",")
        mintCompletionTokens = CInt(Trim(Mid(strJson, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "total_tokens"
    If InStr(1, strJson, """total_tokens"":") > 0 Then
        intStartPos = InStr(1, strJson, """total_tokens"":") + Len("""total_tokens"":")
        intEndPos = InStr(intStartPos, strJson, "}")
        mintTotalTokens = CInt(Trim(Mid(strJson, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "message_role"
    If InStr(1, strJson, """role"":""") > 0 Then
        intStartPos = InStr(1, strJson, """role"":""") + Len("""role"":""")
        intEndPos = InStr(intStartPos, strJson, """")
        mstrMessageRole = Trim(Mid(strJson, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "message_content"
    If InStr(1, strJson, """content"":""") > 0 Then
        intStartPos = InStr(1, strJson, """content"":""") + Len("""content"":""")
        intEndPos = InStr(intStartPos, strJson, """},") ' end position is now before "}," sequence
        strTemp = Mid(strJson, intStartPos, intEndPos - intStartPos)
        strTemp = Replace(strTemp, "\""", """") ' Replace escaped quotes with actual quotes
        mstrMessageContent = Trim(strTemp)
    End If

    
    ' Extract "finish_reason"
    If InStr(1, strJson, """finish_reason"":""") > 0 Then
        intStartPos = InStr(1, strJson, """finish_reason"":""") + Len("""finish_reason"":""")
        intEndPos = InStr(intStartPos, strJson, """")
        mstrFinishReason = Trim(Mid(strJson, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "index"
    If InStr(1, strJson, """index"":") > 0 Then
        intStartPos = InStr(1, strJson, """index"":") + Len("""index"":")
        intEndPos = InStr(intStartPos, strJson, "}")
        mintIndex = CInt(Trim(Mid(strJson, intStartPos, intEndPos - intStartPos)))
    End If
End Sub


Public Sub ParseTextCompletionJSON(ByVal strJson As String)
'Purpose: This method is for parsing OpenAI's strJson from it's text completion end point

    Dim intStartPos As Integer
    Dim intEndPos As Integer
    Dim strTemp As String
    
    ' Extract "id"
    If InStr(1, strJson, """id"":""") > 0 Then
        intStartPos = InStr(1, strJson, """id"":""") + Len("""id"":""")
        intEndPos = InStr(intStartPos, strJson, """")
        mstrId = Trim(Mid(strJson, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "object"
    If InStr(1, strJson, """object"":""") > 0 Then
        intStartPos = InStr(1, strJson, """object"":""") + Len("""object"":""")
        intEndPos = InStr(intStartPos, strJson, """")
        mstrObject = Trim(Mid(strJson, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "created"
    If InStr(1, strJson, """created"":") > 0 Then
        intStartPos = InStr(1, strJson, """created"":") + Len("""created"":")
        intEndPos = InStr(intStartPos, strJson, ",")
        mstrCreated = CLng(Trim(Mid(strJson, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "model"
    If InStr(1, strJson, """model"":""") > 0 Then
        intStartPos = InStr(1, strJson, """model"":""") + Len("""model"":""")
        intEndPos = InStr(intStartPos, strJson, """")
        mstrModel = Trim(Mid(strJson, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "prompt_tokens"
    If InStr(1, strJson, """prompt_tokens"":") > 0 Then
        intStartPos = InStr(1, strJson, """prompt_tokens"":") + Len("""prompt_tokens"":")
        intEndPos = InStr(intStartPos, strJson, ",")
        mintPromptTokens = CInt(Trim(Mid(strJson, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "completion_tokens"
    If InStr(1, strJson, """completion_tokens"":") > 0 Then
        intStartPos = InStr(1, strJson, """completion_tokens"":") + Len("""completion_tokens"":")
        intEndPos = InStr(intStartPos, strJson, ",")
        mintCompletionTokens = CInt(Trim(Mid(strJson, intStartPos, intEndPos - intStartPos)))
    End If
    
    ' Extract "total_tokens"
    If InStr(1, strJson, """total_tokens"":") > 0 Then
        intStartPos = InStr(1, strJson, """total_tokens"":") + Len("""total_tokens"":")
        intEndPos = InStr(intStartPos, strJson, "}")
        mintTotalTokens = CInt(Trim(Mid(strJson, intStartPos, intEndPos - intStartPos)))
    End If
    
        ' Extract "text"
    If InStr(1, strJson, """text"":""") > 0 Then
        intStartPos = InStr(1, strJson, """text"":""") + Len("""text"":""")
        intEndPos = InStr(intStartPos, strJson, """,""") ' end position is now before the sequence ","
        strTemp = Mid(strJson, intStartPos, intEndPos - intStartPos)
        strTemp = Replace(strTemp, "\""", """") ' Replace escaped quotes with actual quotes
        mstrText = Trim(strTemp)
    End If
    
    ' Extract "logprobs"
    If InStr(1, strJson, """logprobs"":") > 0 Then
        intStartPos = InStr(1, strJson, """logprobs"":") + Len("""logprobs"":")
        intEndPos = InStr(intStartPos, strJson, ",") ' end position is now before the sequence ","
        strTemp = Mid(strJson, intStartPos, intEndPos - intStartPos)
        strTemp = Replace(strTemp, "\""", """") ' Replace escaped quotes with actual quotes
        mstrLogprobs = Trim(strTemp)
    End If

    
    ' Extract "finish_reason"
    If InStr(1, strJson, """finish_reason"":""") > 0 Then
        intStartPos = InStr(1, strJson, """finish_reason"":""") + Len("""finish_reason"":""")
        intEndPos = InStr(intStartPos, strJson, """")
        mstrFinishReason = Trim(Mid(strJson, intStartPos, intEndPos - intStartPos))
    End If
    
    ' Extract "index"
    If InStr(1, strJson, """index"":") > 0 Then
        intStartPos = InStr(1, strJson, """index"":") + Len("""index"":")
        intEndPos = InStr(intStartPos, strJson, ",")
        mintIndex = CInt(Trim(Mid(strJson, intStartPos, intEndPos - intStartPos)))
    End If
End Sub


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



