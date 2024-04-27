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
    Contain_
    EndWith_
    EvaluateTo_
    StartWith_
End Enum

Public Enum flHave
    ErrorDescriptionOf_
    ErrorNumberOf_
    ExactSameElementsAs_
    LengthBetween_
    LengthOf_
    MaxLengthOf_
    MinLengthOf_
    SameElementsAs_
    SameTypeAs_
    SameUniqueElementsAs_
End Enum

Public Enum flBe
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
    Something_
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
    Optional LowerVal As Variant, _
    Optional HigherVal As Variant)
End Sub

Public Function Should(shouldType As flShould, Optional testingInput As Variant) As String
End Function

Public Function Be(beType As flBe, Optional TesingInput As String = "", Optional LowerVal As Variant, Optional HigherVal As Variant)
End Function

Public Function Have(haveType As flHave, Optional TesingInput As String = "", Optional LowerVal As Variant, Optional HigherVal As Variant)
End Function

Public Function buildFluentPath(test As ITest, FluentPath As String, tempFluentPath As String) As String
End Function

Public Function getFunctionExpr(test As ITest)
End Function

Public Function getFluentPath(FluentPath As String, test As ITest, Optional notShould As Boolean = False)
End Function

Public Sub addFluentPathNew(ByVal test As ITest, ByVal negateValue As Boolean)
End Sub

Public Function ToString(ByVal nv As Variant) As String
End Function

Public Function DatastructureIsEmpty(v As Variant) As Boolean
End Function

Public Function InputToStringRecur(v As Variant) As String
End Function

Public Function getResultNumber(result As Variant)
End Function

Public Function IsDataStructure(ByVal DataStructure As Variant)
End Function

Public Function getTestDictCounter() As Scripting.Dictionary
End Function

