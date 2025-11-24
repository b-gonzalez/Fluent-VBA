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

Private Function IUtilities_NegateIfBool(ByVal result As Variant) As Variant
    Dim tempResult As Variant
    
    tempResult = result
    
    If Not VBA.Information.IsEmpty(result) And Not VBA.Information.IsNull(result) Then
        tempResult = Not result
    End If
    
    IUtilities_NegateIfBool = tempResult
End Function

Function IUtilities_getElementFromCollection(col As VBA.Collection, index As Long) As VBA.Collection
    Dim indexOffset As Long
    Dim temp As Variant
    Dim tempCol As Collection
    
    Set tempCol = New Collection
    indexOffset = 1
    
    If Abs(index) <= col.Count Then
        If index > 0 Then
            tempCol.Add col(index)
        ElseIf index < 0 Then
            tempCol.Add col(col.Count + indexOffset + index)
        End If
    Else
        tempCol.Add Null
    End If
    
    
    Set IUtilities_getElementFromCollection = tempCol
End Function
