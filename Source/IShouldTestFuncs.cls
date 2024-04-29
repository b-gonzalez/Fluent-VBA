VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShouldTestFuncs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Contain( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal NegateValue As Boolean = False) As Variant
End Function

Public Function StartWith( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal NegateValue As Boolean = False) As Variant
End Function

Public Function EndWith( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal NegateValue As Boolean = False) As Variant
End Function

Public Function EvaluateTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal NegateValue = False) As Variant
End Function

