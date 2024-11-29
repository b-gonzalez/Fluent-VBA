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

Private pCleanTestValueStr As Boolean
Private pCleanTestInputStr As Boolean
Private pCleanTestStrings As Boolean

Private pCleanStringDict As Scripting.Dictionary

Private Sub Class_Initialize()
    pCleanTestValueStr = False
    pCleanTestInputStr = False
    pCleanTestStrings = False

    Set pCleanStringDict = New Scripting.Dictionary
    
    pCleanStringDict.Add " ", ""
    pCleanStringDict.Add """", ""
    
End Sub

Private Property Let ITestStrings_CleanTestValueStr(ByVal value As Boolean)
    pCleanTestValueStr = value
End Property

Private Property Get ITestStrings_CleanTestValueStr() As Boolean
    ITestStrings_CleanTestValueStr = pCleanTestValueStr
End Property

Private Property Let ITestStrings_CleanTestInputStr(ByVal value As Boolean)
    pCleanTestInputStr = value
End Property

Private Property Get ITestStrings_CleanTestInputStr() As Boolean
    ITestStrings_CleanTestInputStr = pCleanTestInputStr
End Property

Private Property Let ITestStrings_CleanTestStrings(ByVal value As Boolean)
    pCleanTestStrings = value
    
    pCleanTestValueStr = value
    pCleanTestInputStr = value
End Property

Private Property Get ITestStrings_CleanTestStrings() As Boolean
    ITestStrings_CleanTestStrings = pCleanTestStrings
End Property

Private Property Get ITestStrings_CleanStringDict() As Scripting.Dictionary
    Set ITestStrings_CleanStringDict = pCleanStringDict
End Property

Private Sub ITestStrings_AddToCleanStringDict(ByVal key As String, Optional ByVal value As String = "", Optional ByVal replaceIfExists As Boolean = False)
    If replaceIfExists Then
        pCleanStringDict(key) = value
    Else
        If Not pCleanStringDict.Exists(key) Then
            pCleanStringDict.Add key, value
        End If
    End If
End Sub

Private Function ITestStrings_CleanString(ByVal TestStr As Variant, Optional ByVal CleanDict As Scripting.Dictionary = Nothing) As String
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
