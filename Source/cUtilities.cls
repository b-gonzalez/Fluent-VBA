VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IUtilities

Private pCleanStringDict As Scripting.Dictionary

Public Function IUtilities_CleanString(TestStr As Variant, Optional CleanDict As Scripting.Dictionary = Nothing) As String
    Dim elem As Variant
    Dim tempStr As Variant

    tempStr = TestStr
    
    If CleanDict Is Nothing Then
        Set CleanDict = pCleanStringDict
    End If

    For Each elem In CleanDict.Keys
        tempStr = Replace(tempStr, elem, CleanDict(elem))
    Next elem

    IUtilities_CleanString = tempStr
End Function

Public Property Set IUtilities_CleanStringDict(value As Scripting.Dictionary)
    Set pCleanStringDict = value
End Property

Private Function IUtilities_NegateIfBool(Result As Variant) As Variant
    Dim tempResult As Variant
    
    tempResult = Result
    
    If Not Information.IsEmpty(Result) And Not Information.IsNull(Result) Then
        tempResult = Not Result
    End If
    
    IUtilities_NegateIfBool = tempResult
End Function
