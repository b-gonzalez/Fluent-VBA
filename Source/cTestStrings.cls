VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTestStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ITestStrings

Private pCleanTestValStr As Boolean
Private pCleanTestInputStr As Boolean
Private pCleanTestStrings As Boolean

Private pCleanStringDict As Scripting.Dictionary

Private Sub Class_Initialize()
    pCleanTestValStr = False
    pCleanTestInputStr = False
    pCleanTestStrings = False

    Set pCleanStringDict = New Scripting.Dictionary
    
    pCleanStringDict.Add " ", ""
    pCleanStringDict.Add """", ""
    
End Sub

Private Property Let ITestStrings_CleanTestValStr(value As Boolean)
    pCleanTestValStr = value
End Property

Private Property Get ITestStrings_CleanTestValStr() As Boolean
    ITestStrings_CleanTestValStr = pCleanTestValStr
End Property

Private Property Let ITestStrings_CleanTestInputStr(value As Boolean)
    pCleanTestInputStr = value
End Property

Private Property Get ITestStrings_CleanTestInputStr() As Boolean
    ITestStrings_CleanTestInputStr = pCleanTestInputStr
End Property

Private Property Let ITestStrings_CleanTestStrings(value As Boolean)
    pCleanTestStrings = value
    
    pCleanTestValStr = value
    pCleanTestInputStr = value
End Property

Private Property Get ITestStrings_CleanTestStrings() As Boolean
    ITestStrings_CleanTestStrings = pCleanTestStrings
End Property

Private Property Get ITestStrings_CleanStringDict() As Scripting.Dictionary
    Set ITestStrings_CleanStringDict = pCleanStringDict
End Property

Private Sub ITestStrings_AddToCleanStringDict(key As String, Optional value As String = "", Optional replaceIfExists As Boolean = False)
    If replaceIfExists Then
        pCleanStringDict(key) = value
    Else
        If Not pCleanStringDict.Exists(key) Then
            pCleanStringDict.Add key, value
        End If
    End If
End Sub

Private Function ITestStrings_CleanString(TestStr As Variant, Optional CleanDict As Scripting.Dictionary = Nothing) As String
    Dim elem As Variant
    Dim tempStr As Variant

    tempStr = TestStr
    
    If CleanDict Is Nothing Then
        Set CleanDict = pCleanStringDict
    End If

    For Each elem In CleanDict.Keys
        tempStr = Replace(tempStr, elem, CleanDict(elem))
    Next elem

    ITestStrings_CleanString = tempStr
End Function
