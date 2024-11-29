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
    
    If Not Information.IsEmpty(result) And Not Information.IsNull(result) Then
        tempResult = Not result
    End If
    
    IUtilities_NegateIfBool = tempResult
End Function
