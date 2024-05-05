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
    [_first] = 0
    Contain_
    EndWith_
    EvaluateTo_
    StartWith_ = 3
    [_last] = 100
End Enum

Public Enum flHave
    [_first] = 101
    ErrorDescriptionOf_
    ErrorNumberOf_
    ExactSameElementsAs_
    LengthBetween_
    LengthOf_
    MaxLengthOf_
    MinLengthOf_
    SameElementsAs_
    SameTypeAs_
    SameUniqueElementsAs_ = 111
    [_last] = 200
End Enum

Public Enum flBe
    [_first] = 201
    Alphabetic_
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
    Something_ = 216
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
    ByVal Result As Variant, _
    ByVal NegateValue As Boolean, _
    ByVal testingValue As Variant, _
    Optional ByVal testingInput As Variant, _
    Optional ByVal LowerVal As Variant, _
    Optional ByVal HigherVal As Variant)
End Sub

Public Function Should(ByVal shouldType As flShould, Optional ByVal testingInput As Variant) As String
End Function

Public Function Be(ByVal beType As flBe, _
Optional ByVal TesingInput As String = "", _
Optional ByVal LowerVal As Variant, _
Optional ByVal HigherVal As Variant)
End Function

Public Function Have(ByVal haveType As flHave, _
Optional ByVal TesingInput As String = "", _
Optional ByVal LowerVal As Variant, _
Optional ByVal HigherVal As Variant)
End Function

Public Function buildFluentPath(ByVal test As ITest, ByVal FluentPath As String, ByVal tempFluentPath As String) As String
End Function

Public Function getFunctionExpr(ByVal test As ITest)
End Function

Public Function getFluentPath(ByVal FluentPath As String, ByVal test As ITest, Optional ByVal notShould As Boolean = False)
End Function

Public Sub addFluentPathNew(ByVal test As ITest, ByVal NegateValue As Boolean)
End Sub

Public Function ToString(ByVal nv As Variant) As String
End Function

Public Function DatastructureIsEmpty(ByVal v As Variant) As Boolean
End Function

Public Function InputToStringRecur(ByVal v As Variant) As String
End Function

Public Function getResultNumber(ByVal Result As Variant)
End Function

Public Function IsDataStructure(ByVal DataStructure As Variant)
End Function

Public Function getTestDictCounter() As Scripting.Dictionary
End Function

Private Function datastructureInCollection(ByVal dsName As String, ByVal col As Collection)
End Function
