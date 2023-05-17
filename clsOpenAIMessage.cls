VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenAIMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Project: OpenAI VBA Framework
' Class:  clsOpenAIMessage
' Description: Wrapper for a single message
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

Private mstrRole As String
Private mstrContent As String

Private mcollMessageParts As Collection

Private Function IOpenAINameProvider_GetClassName() As String
    IOpenAINameProvider_GetClassName = "clsOpenAIMessage"
End Function

Private Function IOpenAINameProvider_ToString() As String
    IOpenAINameProvider_ToString = GetMessageContentString()
End Function

Private Sub Class_Initialize()
    Set mcollMessageParts = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcollMessageParts = Nothing
End Sub


Public Sub Add(ByVal strKeyName As String, strKey As String, ByVal strValueName As String, ByVal strValue As String)
'Purpose: Main access point for adding message to the class

    Dim strPart As String
    
    strPart = FormatMessage(strKeyName, strKey, strValueName, strValue)

    If Len(strPart) > 0 Then
        If Not mcollMessageParts Is Nothing Then
            mcollMessageParts.Add strPart
        End If
    End If
    
End Sub


Private Function FormatMessage(ByVal strKeyName As String, ByVal strKey As String, ByVal strValueName As String, ByVal strValue As String) As String
'Purpose: Format the message part into key/value pairs

    If (Len(strKeyName) > 0) And (Len(strKey) > 0) And (Len(strValueName) > 0) And (Len(strValue) > 0) Then
        FormatMessage = """" & strKeyName & """: """ & strKey & """, """ & strValueName & """: """ & strValue & """"
    End If
End Function


Public Function GetMessageContentString() As String
'Purpose: Joins in all the message parts into a python dictionary like string

    Dim strOutput As String
    Dim strPart As Variant
    
    strOutput = "{"
    
    For Each strPart In mcollMessageParts
        strOutput = strOutput & strPart & ", "
    Next strPart
    
    ' Remove the trailing comma and space
    If mcollMessageParts.Count > 0 Then
        strOutput = Left(strOutput, Len(strOutput) - 2)
    End If
    
    strOutput = strOutput & "}"
    
    GetMessageContentString = strOutput
    
End Function

