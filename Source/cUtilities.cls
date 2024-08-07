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

Private Function IUtilities_NegateIfBool(Result As Variant) As Variant
    Dim tempResult As Variant
    
    tempResult = Result
    
    If Not Information.IsEmpty(Result) And Not Information.IsNull(Result) Then
        tempResult = Not Result
    End If
    
    IUtilities_NegateIfBool = tempResult
End Function
