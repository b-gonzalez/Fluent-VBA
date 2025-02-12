VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExpressions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum flShould
    [_first] = 1
    Contain_ = [_first]
    EndWith_
    EvaluateTo_
    StartWith_ '4
    [_last] = 100
End Enum

Public Enum flHave
    [_first] = 101
    DepthCountOf_ = [_first]
    Elements_
    ElementsInDataStructure_
    ErrorDescriptionOf_
    ErrorNumberOf_
    ExactSameElementsAs_
    LengthBetween_
    LengthOf_
    MaxLengthOf_
    MinLengthOf_
    NestedCountOf_
    Procedure_
    SameElementsAs_
    SameTypeAs_
    SameUniqueElementsAs_ '115
    [_last] = 200
End Enum

Public Enum flBe
    [_first] = 201
    Alphabetic_ = [_first]
    Alphanumeric_
    Between_
    Erroneous_
    EqualTo_
    GreaterThan_
    GreaterThanOrEqualTo_
    IdenticalTo_
    InDataStructure_
    InDataStructures_
    LessThan_
    LessThanOrEqualTo_
    Numeric_
    OneOf_
    Something_ '215
    [_last] = 300
End Enum

Public Property Set setPrinting(ByVal value As cPrinting)
End Property

Public Property Set setTests(ByVal value As cTests)
End Property

Public Property Set setMeta(ByVal value As cMeta)
End Property

Public Sub addToTestAndFluentPath( _
    ByVal functionName As String, _
    ByVal result As Variant, _
    ByVal negateValue As Boolean, _
    ByVal testingValue As Variant, _
    Optional ByVal testingInput As Variant, _
    Optional ByVal lowerVal As Variant, _
    Optional ByVal higherVal As Variant, _
    Optional ByVal cleanedTestValue As Variant, _
    Optional ByVal cleanedTestInput As Variant, _
    Optional ByVal Actual As Variant, _
    Optional ByVal recurIterFunc As Variant, _
    Optional ByVal TestingValueIsSelfReferential As Variant, _
    Optional ByVal TestingInputIsSelfReferential As Variant _
)
End Sub

Public Function Should(ByVal shouldType As flShould, Optional ByVal testingInput As Variant) As String
End Function

Public Function Be(ByVal beType As flBe, _
    Optional ByVal TesingInput As String = "", _
    Optional ByVal lowerVal As Variant, _
    Optional ByVal higherVal As Variant) As String
End Function

Public Function Have(ByVal haveType As flHave, _
    Optional ByVal TesingInput As String = "", _
    Optional ByVal lowerVal As Variant, _
    Optional ByVal higherVal As Variant) As String
End Function

Public Function buildFluentPath(ByVal test As ITest, ByVal FluentPath As String, ByVal tempFluentPath As String) As String
End Function

Public Function getFunctionExpr(ByVal test As ITest) As String
End Function

Public Function getFluentPath( _
    ByVal FluentPath As String, _
    ByVal test As ITest, _
    Optional ByVal notShould As Boolean = False) As String
End Function

Public Sub addFluentPathNew(ByVal test As ITest, ByVal negateValue As Boolean)
End Sub

Public Function InputToString(ByVal nv As Variant, Optional ByVal algo As Variant) As String
End Function

Public Function DatastructureIsEmpty(ByVal v As Variant) As Boolean
End Function

Public Function inputToStringRecur(ByVal v As Variant) As String
End Function

Public Function getResultNumber(ByVal result As Variant) As Long
End Function

Public Function inputToStringIter(ByVal v As Variant) As String
End Function

Private Function getNestedElementsFromCollection(ByVal col As VBA.Collection) As VBA.Collection
End Function

Public Function getDsCount(ByVal dataStructure As Variant) As Long
End Function

Private Function getDsType(ByVal v As Variant) As String
End Function
