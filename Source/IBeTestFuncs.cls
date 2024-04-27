VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBeTestFuncs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InDataStructures( _
ByVal testingValue As Variant, _
ByVal method As flAlgorithm, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant) As Variant
End Function

Public Function InDataStructure( _
ByVal testingValue As Variant, _
ByVal DataStructure As Variant, _
Optional ByVal method As flAlgorithm = flRecursive, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
End Function

Private Function getNestedElementsRecur(ByVal DataStructure As Variant, Optional ByRef col As Collection = Nothing) As Collection
End Function

Private Function getNestedElementsIter(ByVal DataStructure As Variant) As Collection
End Function

Public Function GreaterThan(ByVal testingValue As Variant, ByVal testingInput As Double, Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function GreaterThanOrEqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
End Function

Public Function LessThan( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function LessThanOrEqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
End Function

Public Function EqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
End Function

Public Function Something(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function Between( _
ByVal testingValue As Variant, _
ByVal LowerVal As Double, _
ByVal HigherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function OneOf(ByVal testingValue As Variant, ByVal negateValue As Boolean, ByVal testingInput As Variant) As Variant
End Function

Public Function Alphabetic(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function Numeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function Alphanumeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function Erroneous(ByVal testingValue As Variant, Optional ByVal negateValue = False) As Variant
End Function

Public Function IdenticalTo( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant, _
ByVal updateFluentPath As Boolean) As Variant
End Function

