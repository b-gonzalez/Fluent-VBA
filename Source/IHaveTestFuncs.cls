VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHaveTestFuncs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function LengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal NegateValue As Boolean = False) As Variant
End Function

Public Function MaxLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal NegateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
End Function

Public Function MinLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal NegateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
End Function

Public Function LengthBetween( _
ByVal testingValue As Variant, _
ByVal LowerVal As Double, _
ByVal HigherVal As Double, _
Optional ByVal NegateValue As Boolean = False) As Variant
End Function

Public Function ErrorNumberOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal NegateValue = False) As Variant
End Function

Public Function ErrorDescriptionOf( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal NegateValue = False) As Variant
End Function

Public Function SameTypeAs( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal NegateValue = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
End Function

Public Function ExactSameElementsAs( _
ByVal testingValue As Variant, _
ByVal NegateValue As Boolean, _
ByVal updateFluentPath As Boolean, _
ByVal testingInput As Variant) As Variant
End Function

Public Function SameUniqueElementsAs( _
ByVal testingValue As Variant, _
ByVal NegateValue As Boolean, _
ByVal testingInput As Variant) As Variant
End Function

Public Function SameElementsAs( _
ByVal testingValue As Variant, _
ByVal NegateValue As Boolean, _
ByVal testingInput As Variant) As Variant
End Function


