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
    Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function MaxLengthOf( _
    ByVal testingValue As Variant, _
    ByVal testingInput As Long, _
    Optional ByVal negateValue As Boolean = False, _
    Optional ByVal updateFluentPath As Boolean = False) As Variant
End Function

Public Function MinLengthOf( _
    ByVal testingValue As Variant, _
    ByVal testingInput As Long, _
    Optional ByVal negateValue As Boolean = False, _
    Optional ByVal updateFluentPath As Boolean = False) As Variant
End Function

Public Function LengthBetween( _
    ByVal testingValue As Variant, _
    ByVal lowerVal As Double, _
    ByVal higherVal As Double, _
    Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function ErrorNumberOf( _
    ByVal testingValue As Variant, _
    ByVal testingInput As Long, _
    Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function ErrorDescriptionOf( _
    ByVal testingValue As Variant, _
    ByVal testingInput As String, _
    Optional ByVal negateValue As Boolean = False) As Variant
End Function

Public Function SameTypeAs( _
    ByVal testingValue As Variant, _
    ByVal testingInput As Variant, _
    Optional ByVal negateValue As Boolean = False, _
    Optional ByVal updateFluentPath As Boolean = False) As Variant
End Function

Public Function ExactSameElementsAs( _
    ByVal testingValue As Variant, _
    ByVal negateValue As Boolean, _
    ByVal updateFluentPath As Boolean, _
    ByVal testingInput As Variant) As Variant
End Function

Public Function SameUniqueElementsAs( _
    ByVal testingValue As Variant, _
    ByVal negateValue As Boolean, _
    ByVal testingInput As Variant) As Variant
End Function

Public Function SameElementsAs( _
    ByVal testingValue As Variant, _
    ByVal negateValue As Boolean, _
    ByVal testingInput As Variant) As Variant
End Function

Public Function Procedure( _
    ByVal testingValue As Variant, _
    ByVal procName As String, _
    ByVal callType As VBA.VbCallType, _
    ByVal negateValue As Boolean _
) As Variant
End Function

Public Function Elements( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Variant, _
    ByVal negateValue As Boolean, _
    ByVal updateFluentPath As Boolean _
    ) As Variant
End Function

Public Function ElementsInDataStructure( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Variant, _
    ByVal negateValue As Boolean _
    ) As Variant
End Function

Public Function DepthCountOf( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Long, _
    ByVal negateValue As Boolean _
    ) As Variant
End Function

Private Function getDepthCountRecur( _
    ByVal testingValue As Variant, _
    Optional ByVal counter As Long = 1, _
    Optional ByRef depthCounter As Long = 0 _
) As Long
End Function

Private Function getDepthCountIter(ByVal ds As Variant) As Long
End Function

Public Function NestedCountOf( _
    ByVal testingValue As Variant, _
    ByRef testingInput As Long, _
    ByVal negateValue As Boolean _
    ) As Variant
End Function

