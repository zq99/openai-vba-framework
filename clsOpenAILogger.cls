VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenAILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Project: OpenAI VBA Framework
' Class:  clsOpenAILogger
' Description: Handles developer messages in the framework for Immediate Window
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
Private mobjClass As IOpenAINameProvider

Private mblnIsMessageRequired As Boolean

Private Function IOpenAINameProvider_GetClassName() As String
    IOpenAINameProvider_GetClassName = "clsOpenAILogger"
End Function

Private Function IOpenAINameProvider_ToString() As String
    IOpenAINameProvider_ToString = "IsMessageRequired=" & Me.IsMessageRequired
End Function

Public Property Let IsMessageRequired(ByVal value As Boolean)
    mblnIsMessageRequired = value
End Property

Public Property Get IsMessageRequired() As Boolean
    IsMessageRequired = mblnIsMessageRequired
End Property

Private Sub Class_Terminate()
    Set mobjClass = Nothing
End Sub

Public Sub SetClass(ByVal Obj As IOpenAINameProvider)
    Set mobjClass = Obj
End Sub

Public Sub PrintMessage(ParamArray vntMessage() As Variant)
' Purpose: Takes in an open ended list of string variables and appends then together as a message to
'          output to the immediate window

    If IsEmpty(vntMessage) Then
        Exit Sub
    Else
        Dim i As Integer
        Dim strAll As String
        'concatenate message elements into into one string
        For i = LBound(vntMessage) To UBound(vntMessage)
            strAll = strAll & IIf(Len(strAll) > 0, " | ", "") & CStr(vntMessage(i))
        Next i
        'output string to immediate window
        If Len(Trim(strAll)) > 0 Then
            Call LogMessage(strAll)
        End If
    End If

End Sub

Private Sub LogMessage(ByVal strMessage As String)
'Purpose: The main logging routine for messages, which can be suppressed by calling routines

    If Not mobjClass Is Nothing Then
        If Me.IsMessageRequired Then
            Debug.Print fncGetDateTimeStamp & vbTab & mobjClass.GetClassName & " : " & strMessage
        End If
    End If
End Sub

Public Sub LogCriticalMessage(ByVal strMessage As String, Optional ByVal blnIsBorderRequired As Boolean = False, Optional ByVal strBorderCharacter As String = "*", Optional ByVal blnAddModuleName As Boolean = True, Optional ByVal strLabel As String = "WARNING")
'Purpose: This method always outputs to the immediate window regardless of whether logging is set to False

    If Not mobjClass Is Nothing Then
        Dim strMsg As String
        Dim strName As String
        
        strName = IIf(blnAddModuleName = True, mobjClass.GetClassName, "")
        strMsg = fncGetDateTimeStamp & vbTab & strName & " " & strLabel & ": " & strMessage
        
        If blnIsBorderRequired Then
            Dim strBorder As String
            strBorder = String(Len(strMsg), strBorderCharacter)
            Debug.Print strBorder
            Debug.Print strMsg
            Debug.Print strBorder
        Else
            Debug.Print strMsg
        End If
        
    End If

End Sub


Public Sub LogVBAError(ByVal objErr As Error)
'Purpose: This method always outputs to the immediate window regardless of whether logging is set to False

    If Not mobjClass Is Nothing Then
        Dim strMsg As String
        Dim strName As String
        
        strName = mobjClass.GetClassName
        strMsg = fncGetDateTimeStamp & vbTab & strName & " VBA ERROR: [" & Err.Number & "]" & vbTab & Err.Description
        Debug.Print strMsg
    End If

End Sub

Private Function fncGetDateTimeStamp() As String
'Purpose: All logged messages are outputted with a time stamp unless caling functions suppress this
    fncGetDateTimeStamp = Format(Now, "yyyy-MM-dd hh:mm:ss")
End Function
