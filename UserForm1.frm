VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8535.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function AverageVar(ByVal variable1 As Integer, ByVal variable2 As Integer) As Integer
    AverageVar = (variable1 + variable2) / 2
End Function

Function GetMax(ByVal value1 As Double, ByVal value2 As Double) As Double
    ' Compare the two values and return the maximum
    If value1 >= value2 Then
        GetMax = value1
    Else
        GetMax = value2
    End If
End Function

Function InnerCompareTemp(originTemp As Integer, destTemp As Integer) As Integer
    Dim innerResult As Integer
    Dim avg As Integer
    avg = AverageVar(originTemp, destTemp)
    
    'refrigerated medication
    If Refrigerate.Value = True Then
        If avg <= 46 Then
            innerResult = 0
        ElseIf avg <= 70 Then
            innerResult = 24
        ElseIf avg > 70 Then
            innerResult = 48
        Else
            MsgBox ("An Error Occured." & vbCrLf & "Please check the zip codes entered.")
        End If
    'room temp medication
    ElseIf Refrigerate.Value = False Then
        If avg <= 46 Then
            innerResult = 0
        ElseIf avg <= 70 Then
            innerResult = 0
        ElseIf avg > 70 Then
            innerResult = 0
        Else
            MsgBox ("An Error Occured." & vbCrLf & "Please check the zip codes entered.")
        End If
    Else
        MsgBox ("An error occured on the form.")
    End If
    
    'return
    InnerCompareTemp = innerResult
End Function

Function OuterCompareTemp(originTemp As Integer, destTemp As Integer) As Integer
    Dim outerResult As Integer
    Dim avg As Integer
    Dim max As Integer
    avg = AverageVar(originTemp, destTemp)
    max = GetMax(originTemp, destTemp)
    
    If Refrigerate.Value = True Then
        If max < 90 Then
            If max >= 46 And avg <= 46 Then
                outerResult = 24
            ElseIf avg <= 46 Then
                outerResult = 0
            ElseIf avg > 46 And avg <= 70 Then
                outerResult = 24
            ElseIf avg > 70 Then
                outerResult = 48
            End If
        ElseIf avg <= 46 Then
            outerResult = 24
        ElseIf avg > 46 And avg <= 70 Then
            outerResult = 48
        ElseIf avg > 70 Then
            outerResult = 72
        End If
    ElseIf Refrigerate.Value = False Then
        If max < 90 Then
            If max >= 46 And avg <= 46 Then
                outerResult = 24
            ElseIf avg <= 46 Then
                outerResult = 0
            ElseIf avg > 46 And avg <= 70 Then
                outerResult = 24
            ElseIf avg > 70 Then
                outerResult = 48
            End If
        ElseIf avg <= 46 Then
            outerResult = 0
        ElseIf avg > 46 And avg <= 70 Then
            outerResult = 24
        ElseIf avg > 70 Then
            outerResult = 48
        End If
    Else
        MsgBox ("An error occured on the form.")
    End If
        
    'return
    OuterCompareTemp = outerResult
    
End Function


Private Sub CommandButton1_Click()
    Dim JsonObject As Object
    Dim objRequest As Object
    Dim weatherUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim originValue As String
    Dim originTemp As Integer
    
    originValue = Origin.Value
    
    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    weatherUrl = "https://api.openweathermap.org/data/2.5/weather?zip=" + originValue + ",US&appid={apiKey}&units=imperial"
    blnAsync = True
     
     With objRequest
        .Open "GET", weatherUrl, blnAsync
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & token
        .Send
        'spin wheels whilst waiting for response
        While objRequest.readyState <> 4
            DoEvents
        Wend
        strResponse = .responseText
    End With
    Set JsonObject = JsonConverter.ParseJson(strResponse)
     
     
     
     Dim destinationValue As String
     
     destinationValue = Destination.Value
     Dim JsonObject2 As Object
     Dim objRequest2 As Object
     Dim weatherUrl2 As String
     Dim strResponse2 As String
     Dim destTemp As Integer
     
     Set objRequest2 = CreateObject("MSXML2.XMLHTTP")
     weatherUrl2 = "https://api.openweathermap.org/data/2.5/weather?zip=" + destinationValue + ",US&appid={apiKey}&units=imperial"
     blnAsync = True
     
     With objRequest2
        .Open "GET", weatherUrl2, blnAsync
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & token
        .Send
        'spin wheels whilst waiting for response
        While objRequest2.readyState <> 4
            DoEvents
        Wend
        strResponse2 = .responseText
     End With
        Set JsonObject2 = JsonConverter.ParseJson(strResponse2)
        
    originTemp = JsonObject("main")("temp")
    destTemp = JsonObject2("main")("temp")
    
    MsgBox (originTemp & " " & destTemp)
    
    Dim innerValue As Integer
    innerValue = InnerCompareTemp(originTemp, destTemp)
    
    Dim outerValue As Integer
    outerValue = OuterCompareTemp(originTemp, destTemp)
    
    Inner.Value = CStr(innerValue)
    Outer.Value = CStr(outerValue)
    
End Sub



Private Sub Label2_Click()

End Sub
