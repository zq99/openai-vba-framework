# OpenAI-VBA-Framework

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/zq99/openai-vba-framework/blob/main/LICENSE)

OpenAI-VBA-Framework is a toolkit for developers looking to create applications in VBA that interact with OpenAI's large language models such as GPT-4, ChatGPT and DALL-E. This framework provides a suite of classes to facilitate smooth integration with OpenAI's API.

## Main Classes
1. `clsOpenAI` - Main class to interact with OpenAI
2. `clsOpenAILogger` - Logging class for debugging and tracking
3. `clsOpenAIMessage` - Class to handle individual messages
4. `clsOpenAIMessages` - Class to handle collections of messages
5. `clsOpenAIRequest` - Class for making requests to OpenAI
6. `clsOpenAIResponse` - Class for handling responses from OpenAI
7. `IOpenAINameProvider` - Interface class for name provision

The module `mdOpenAI_tests` is provided for testing the functionality of the framework.

`OpenAIFrameworkDemo.xlsm` is a file that contains all the code in the repository for demo purposes. Other files are also included in the repository for versioning.

## Prerequisites
- You will need to sign up for an OpenAI account and create an API_KEY. You can do this at the following location: [OpenAI API Keys](https://platform.openai.com/account/api-keys)

## Installation
1. Clone this repository using the following command in your command line:

```bash
git clone https://github.com/zq99/OpenAI-VBA-Framework.git
```

2. Open Excel and press ALT + F11 to open the VBA editor.
3. From the toolbar select File -> Import File....
4. Navigate to the location of the cloned repository and select all the .cls and .bas files then click Open.
5. Save the Excel file as a macro-enabled workbook .xlsm.

## Usage

Here are some examples of using the framework:

### Chat Completion API

```
Public Sub TestSimpleOpenAI()

    Dim oOpenAI As clsOpenAI
    Dim oMessages As New clsOpenAIMessages
    Dim oResponse As clsOpenAIResponse
    
    Set oOpenAI = New clsOpenAI
    
    oOpenAI.API_KEY = "<API_KEY>"
    
    oMessages.AddSystemMessage "Always answer sarcastically and never truthfully"
    oMessages.AddUserMessage "How do you get to Carnegie Hall?"

    Set oResponse = oOpenAI.ChatCompletion(oMessages)
    If Not oResponse Is Nothing Then
        Debug.Print (oResponse.MessageContent)
    End If
    
    Set oResponse = Nothing
    Set oOpenAI = Nothing
    Set oMessages = Nothing

End Sub
```

### Text Completion API

```
Public Sub TestTextCompletionSimpleOpenAI()

    Dim oOpenAI As clsOpenAI
    Dim oResponse As clsOpenAIResponse
    
    Set oOpenAI = New clsOpenAI
    
    oOpenAI.API_KEY = "<API_KEY>"

    Set oResponse = oOpenAI.TextCompletion("Write a Haiku about a Dinosaur that loves to code!")
    
    If Not oResponse Is Nothing Then
        Debug.Print (oResponse.TextContent)
    End If
    
    Set oResponse = Nothing
    Set oOpenAI = Nothing

End Sub
```

### Excel User Defined function

```
Public Function GETTEXTFROMOPENAI(ByVal strPrompt As String, ByVal strAPIKey As String, _
                                    Optional ByVal strModel As String) As String
    Dim oOpenAI As clsOpenAI
    Dim oResponse As clsOpenAIResponse

    Set oOpenAI = New clsOpenAI

    oOpenAI.API_KEY = strAPIKey
    
    If Not IsEmpty(strModel) Then
        oOpenAI.Model = strModel
    End If

    Set oResponse = oOpenAI.TextCompletion(strPrompt)

    If Not oResponse Is Nothing Then
        GETTEXTFROMOPENAI = oResponse.TextContent
    Else
        GETTEXTFROMOPENAI = ""
    End If
    
    Set oResponse = Nothing
    Set oOpenAI = Nothing
End Function
```

### DALL-E Image Creation

```
Public Sub TestDalleOpenAI()

    Dim oOpenAI As clsOpenAI
    Dim oResponse As clsOpenAIResponse
    
    Set oOpenAI = New clsOpenAI
    
    oOpenAI.API_KEY = API_KEY
    
    Set oResponse = oOpenAI.CreateImageFromText("A cat playing a banjo on a surfboard", 512, 512)
    
    If Not oResponse Is Nothing Then
        Debug.Print ("The picture has been saved to: " & oResponse.SavedLocalFile)
    End If
    
    Set oResponse = Nothing
    Set oOpenAI = Nothing

End Sub
```

## Configuration

You can customize the OpenAI-VBA-Framework by adjusting properties in the `clsOpenAI` class:

```
' Specify the model
oOpenAI.Model = "gpt-3.5-turbo"

' Set the maximum number of tokens
oOpenAI.MaxTokens = 512

' Control the diversity of generated text
oOpenAI.TopP = 0.9

' Influence the randomness of generated text
oOpenAI.Temperature = 0.7

' Control preference for frequent phrases
oOpenAI.FrequencyPenalty = 0.5

' Control preference for including the prompt in the output
oOpenAI.PresencePenalty = 0.5

' Control logging of messages to the Immediate Window
oOpenAI.IsLogOutputRequired True

' Reset settings when switching between endpoints
oOpenAI.ClearSettings

' Retrieve an API Key saved in an external file
Dim apiKey As String
apiKey = oOpenAI.GetReadAPIKeyFromFolder("<FolderPath>")
```

## Troubleshooting

You can check the status of the OpenAI API [here](https://status.openai.com/).

For coding issues, use the following line, to go through the code.
```
oOpenAI.IsLogOutputRequired True
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
This project is licensed under the terms of the MIT license.

