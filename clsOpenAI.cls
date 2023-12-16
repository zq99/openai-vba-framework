VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------
' Project: OpenAI VBA Framework
' Class:  clsOpenAI
' Description: Main class that controls the framework
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

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Implements IOpenAINameProvider

Private mstrAPI_KEY As String
Private mobjHttpRequest As Object
Private mobjLogger As clsOpenAILogger
Private mobjRequest As clsOpenAIRequest
Private mlngCallsToAPICount As Long

'OpenAI API Endpoints
Private Const API_ENDPOINT_CHAT As String = "https://api.openai.com/v1/chat/completions"
Private Const API_ENDPOINT_IMAGE_CREATION As String = "https://api.openai.com/v1/images/generations"

Private Const DEFAULT_LOCAL_LOCATION As String = "C:\Users\Public\Downloads\"
Private mstrFolderToSave As String

'More models can be found here: https://platform.openai.com/docs/models/overview
Private Const DEFAULT_CHAT_MODEL As String = "gpt-3.5-turbo"
Private Const DEFAULT_CHAT_TOKENS_COUNT As Integer = 512

'Project constants
Private Const UNASSIGNED_VALUE As Integer = -1
Private Const MESSAGE_INVALID_API_KEY As String = "An OpenAI API key is either invalid or has not been specified!"
Private Const HTTP_STATUS_OK As Long = 200 ' OK
Private Const HTTP_REQUEST_COMPLETED As Integer = 4

'This allows configuration of different HTTP Requests
Private Const MSXML_XML As String = "MSXML2.XMLHTTP"
Private Const MSXML_SERVER_XML As String = "MSXML2.ServerXMLHTTP"
Private Const MSXML_DEFAULT As String = MSXML_XML
Private mstrMSXMLType As String

Private Function IOpenAINameProvider_GetClassName() As String
    IOpenAINameProvider_GetClassName = "clsOpenAI"
End Function

Private Function IOpenAINameProvider_ToString() As String
    IOpenAINameProvider_ToString = "Key=" & Me.API_KEY
End Function

Public Property Let API_KEY(ByVal value As String)
    mstrAPI_KEY = value
End Property

Public Property Get API_KEY() As String
    API_KEY = mstrAPI_KEY
End Property

Public Property Let MSXMLType(ByVal value As String)
'This allows calling proceedures to change the default type of XML HTTP Request
    
    'These are the only values allowed for this
    If (value <> Me.MSXML_SERVER_XML_VALUE) And (value <> Me.MSXML_XML_VALUE) Then
        Call mobjLogger.PrintCriticalMessage("Invalid MSXML type specified!")
    Else
        mstrMSXMLType = value
    End If
End Property

Public Property Get MSXMLType() As String
    MSXMLType = mstrMSXMLType
End Property

'This method allows for the MSXML_XML constant to be accessible outside of the class
Public Property Get MSXML_XML_VALUE() As String
    MSXML_XML_VALUE = MSXML_XML
End Property

'This method allows for the MSXML_SERVER_XML constant to be accessible outside of the class
Public Property Get MSXML_SERVER_XML_VALUE() As String
    MSXML_SERVER_XML_VALUE = MSXML_SERVER_XML
End Property

Public Property Let Model(ByVal value As String)
     mobjRequest.Model = value
End Property

Public Property Get Model() As String
     Model = mobjRequest.Model
End Property

Public Property Get CallsToAPICount() As Long
    CallsToAPICount = mlngCallsToAPICount
End Property

Public Property Let MaxTokens(ByVal value As Long)
    mobjRequest.MaxTokens = value
End Property

Public Property Let TopP(ByVal value As Double)
    mobjRequest.TopP = value
End Property

Public Property Let Temperature(ByVal value As Double)
    If (value < 0) Or (value > 1) Then
        Call mobjLogger.PrintCriticalMessage("Temperature setting must be between 0 and 1!", blnAddModuleName:=False)
    End If
    mobjRequest.Temperature = value
End Property

Public Property Get Temperature() As Double
    Temperature = mobjRequest.Temperature
End Property

Public Property Let FrequencyPenalty(ByVal value As Double)
    mobjRequest.FrequencyPenalty = value
End Property

Public Property Let PresencePenalty(ByVal value As Double)
    mobjRequest.PresencePenalty = value
End Property

Public Property Let FolderToSaveTo(ByVal value As String)
    mstrFolderToSave = value
End Property

Public Property Get FolderToSaveTo() As String
    FolderToSaveTo = mstrFolderToSave
End Property

Public Sub Pause(Optional ByVal lngMilliSeconds As Long = 1000)
    mobjLogger.PrintMessage ("Pausing for " & CStr(lngMilliSeconds) & " milliseconds")
    Sleep lngMilliSeconds
End Sub

Public Sub IsLogOutputRequired(ByVal value As Boolean)
'Purpose: Calling routines can switch off messages in this framework from appearing in the Immediate window

    If Not mobjLogger Is Nothing Then
        mobjLogger.IsMessageRequired = value
    End If

End Sub


Public Sub Log(ByVal strMessage As String)
'Purpose: Easy routine to log messages

    If Not mobjLogger Is Nothing Then
        mobjLogger.PrintMessage strMessage
    End If
End Sub


Private Function GetResponseFromAPI(ByVal strRequestJson As String, ByVal strEndPoint As String, Optional ByVal strLocalPath As String = DEFAULT_LOCAL_LOCATION) As clsOpenAIResponse
'Purpose: This handles the request to OpenAI's API URL

    Dim strResponseJson As String
    Dim oResponse As clsOpenAIResponse
    
On Error GoTo ERR_HANDLER:

    'default return value
    Set GetResponseFromAPI = Nothing
    
    Set mobjHttpRequest = CreateObject(mstrMSXMLType)
    
    'talk to OpenAI
    With mobjHttpRequest
        
        If mstrMSXMLType = MSXML_SERVER_XML Then
            .setTimeouts mobjRequest.TimeoutResolve, mobjRequest.TimeoutConnect, _
                         mobjRequest.TimeoutSend, mobjRequest.TimeoutReceive
        End If
        
        .Open "POST", strEndPoint, False
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Authorization", "Bearer " & mstrAPI_KEY
        .Send (strRequestJson)
    
    End With
    
    ' unblock other processes if still querying OpenAI
    Do While mobjHttpRequest.readyState <> HTTP_REQUEST_COMPLETED
        DoEvents
    Loop

    If mobjHttpRequest.Status = HTTP_STATUS_OK Then
    
        mlngCallsToAPICount = mlngCallsToAPICount + 1
    
        'get the json result from the successful request
        strResponseJson = Trim(mobjHttpRequest.ResponseText)
        Log strResponseJson
        
        'format the json result according to which api endpoint used
        If strEndPoint = API_ENDPOINT_CHAT Then
            
            'ChatGPT and GPT4
            Set oResponse = New clsOpenAIResponse
            oResponse.ParseChatJSON (strResponseJson)
            Set GetResponseFromAPI = oResponse
        
        ElseIf strEndPoint = API_ENDPOINT_IMAGE_CREATION Then
        
            'DALL-E image generator
            Set GetResponseFromAPI = GetResponseObjectForImageParse(strResponseJson)
        
        End If
    Else
        mobjLogger.PrintCriticalMessage ("Failed to retrieve data from OpenAI. Response code is " & mobjHttpRequest.Status)
    End If
    
EXIT_HERE:
    Set oResponse = Nothing
    Exit Function
    
ERR_HANDLER:
    mobjLogger.PrintVBAError Err
    GoTo EXIT_HERE
End Function


Private Function GetResponseObjectForImageParse(ByVal strResponseJson As String) As clsOpenAIResponse
'Purpose: this takes the response Json from the OpenAI api, and saves the information as a picture
'         on the local PC, it then returns a response object with a reference to the picture created
    
    Set GetResponseObjectForImageParse = Nothing
    
    Dim oResponse As clsOpenAIResponse
    Dim strImageUrl As String
    
    Set oResponse = New clsOpenAIResponse
    strImageUrl = oResponse.GetImageURLFromImageCreationJSON(strResponseJson)
    
    If Len(strImageUrl) > 0 Then

        Dim strFileName As String
        strFileName = oResponse.GetFileNameFromImageURL(strImageUrl)
        
        If Len(strFileName) > 0 Then
        
            Dim strFullName As String
            strFullName = mstrFolderToSave & strFileName
            
            If Not mobjHttpRequest Is Nothing Then
            
                mobjHttpRequest.Open "GET", strImageUrl, False
                mobjHttpRequest.Send
                
                'convert the byte array to a saved image file

                Dim objStream As Object
                Set objStream = CreateObject("ADODB.Stream")
                
                If Not objStream Is Nothing Then
                    objStream.Open
                    objStream.Type = 1
                    objStream.write mobjHttpRequest.ResponseBody
                    objStream.SaveToFile strFullName
                    objStream.Close
                    Set objStream = Nothing
                    
                    If Len(Dir(strFullName)) > 0 Then
                       oResponse.SavedLocalFile = strFullName
                    End If
                    Set GetResponseObjectForImageParse = oResponse
                End If
            End If
        End If
    End If
    Set oResponse = Nothing
End Function


Private Function IsAPIKeyValid() As Boolean
'Purpose: Check a valid API key has been assigned

    IsAPIKeyValid = IIf(Trim(mstrAPI_KEY) = Empty, False, True)

End Function


Public Function ChatCompletion(ByVal oMessages As clsOpenAIMessages) As clsOpenAIResponse
'Purpose: This is for OpenAI's ChatGPT and GPT4 API

    Set ChatCompletion = Nothing

    If Not IsAPIKeyValid Then
        mobjLogger.PrintCriticalMessage MESSAGE_INVALID_API_KEY, True
        Exit Function
    End If

    If oMessages Is Nothing Then
        Exit Function
    End If
    
    Set mobjRequest.messages = oMessages
    
    If mobjRequest.Model = Empty Then
        mobjRequest.Model = DEFAULT_CHAT_MODEL
    End If
    
    If mobjRequest.MaxTokens = UNASSIGNED_VALUE Then
        mobjRequest.MaxTokens = DEFAULT_CHAT_TOKENS_COUNT
    End If
    
    Log mobjRequest.GetChatSendToAPIJsonString
    
    Set ChatCompletion = GetResponseFromAPI(mobjRequest.GetChatSendToAPIJsonString, API_ENDPOINT_CHAT)
    
End Function


Private Sub Class_Initialize()
    
    mstrMSXMLType = MSXML_DEFAULT
    Set mobjRequest = GetDefaultRequestSettings

    Set mobjLogger = New clsOpenAILogger
    mobjLogger.IsMessageRequired = False
    mobjLogger.SetClass Me
    
    mstrFolderToSave = DEFAULT_LOCAL_LOCATION
    mlngCallsToAPICount = 0
    
    mstrAPI_KEY = Empty

End Sub


Private Sub Class_Terminate()
    Set mobjHttpRequest = Nothing
    Set mobjLogger = Nothing
    Set mobjRequest = Nothing
End Sub


Private Function GetDefaultRequestSettings() As clsOpenAIRequest
'Purpose: These are initial settings for the OpenAI request

    Dim oRequest As clsOpenAIRequest
    Set oRequest = New clsOpenAIRequest
    
    With oRequest
        .Model = Empty
        .MaxTokens = UNASSIGNED_VALUE
        .TopP = 1
        .Temperature = 0.5
        .FrequencyPenalty = 0
        .PresencePenalty = 0
        .ImageHeight = 256
        .ImageWidth = 256
        .TimeoutConnect = 30000
        .TimeoutReceive = 30000
        .TimeoutResolve = 30000
        .TimeoutSend = 60000
    End With
    Set GetDefaultRequestSettings = oRequest
    mlngCallsToAPICount = 0
    
    Set oRequest = Nothing
End Function


Public Sub SetTimeOutDefaults(ByVal lngConnect As Long, ByVal lngReceive As Long, ByVal lngResolve As Long, ByVal lngSend As Long)
    If Not mobjRequest Is Nothing Then
        mobjRequest.TimeoutConnect = lngConnect
        mobjRequest.TimeoutReceive = lngReceive
        mobjRequest.TimeoutResolve = lngResolve
        mobjRequest.TimeoutSend = lngSend
    End If
End Sub


Public Sub ClearSettings()
'Purpose: Reset the settings if switching between endpoints

    Set mobjRequest = GetDefaultRequestSettings

End Sub


Public Function GetReadAPIKeyFromFolder(ByVal strfolderPath As String, Optional ByVal strDefaultTextFileName As String = "apikey") As String
'Purpose: Allows retrieval of an API KEY saved in an external file (possibly stored on a drive only the current user can access)

    Dim intFileNumber As Integer
    Dim strAPIKey As String
    Dim strFilePath As String

    strFilePath = strfolderPath & "\" & strDefaultTextFileName & ".txt" ' Construct the full file path
    
    'check file exists
    If Dir(strFilePath) = "" Then
        GetReadAPIKeyFromFolder = Empty
        Exit Function
    End If

    intFileNumber = FreeFile ' Get the first available file number

    ' Open the file in Input mode. Trappable error occurs if file does not exist.
    Open strFilePath For Input As intFileNumber

        ' Read the contents of the file into the variable.
        Input #intFileNumber, strAPIKey

    ' Close the file.
    Close intFileNumber

    ' Return the API Key
    GetReadAPIKeyFromFolder = strAPIKey
End Function


Public Function CreateImageFromText(ByVal strPrompt As String, ByVal lngWidth As Long, ByVal lngHeight As Long) As clsOpenAIResponse
'Purpose: This is for OpenAI's image creation

    Set CreateImageFromText = Nothing

    If Not IsAPIKeyValid Then
        mobjLogger.PrintCriticalMessage MESSAGE_INVALID_API_KEY, True
        Exit Function
    End If
    
    If strPrompt = Empty Then
        Exit Function
    End If
    
    mobjRequest.prompt = strPrompt
    mobjRequest.ImageHeight = lngHeight
    mobjRequest.ImageWidth = lngWidth
    
    Log mobjRequest.GetDalleImageSendToAPIJsonString
    
    Set CreateImageFromText = GetResponseFromAPI(mobjRequest.GetDalleImageSendToAPIJsonString, API_ENDPOINT_IMAGE_CREATION, DEFAULT_LOCAL_LOCATION)

End Function

