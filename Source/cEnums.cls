VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TestProperty
    TestCategory = 1
    TestInfo
    TestName
    TestResultNumber
End Enum

Public Enum flAlgorithm
    flRecursive = 2 ^ 0
    flIterative = 2 ^ 1
End Enum

Public Enum flTestType
    flIsArray = 2 ^ 0
    flIsNotArray = 2 ^ 1
    flIsDataStructure = 2 ^ 2
    flIsErrObject = 2 ^ 3
    flIsNotErrObject = 2 ^ 4
    flIsNotDataStructure = 2 ^ 5
    flIsNull = 2 ^ 6
    flIsNotNull = 2 ^ 7
    flIsNumeric = 2 ^ 8
    flIsNotNumeric = 2 ^ 9
    flIsObject = 2 ^ 10
    flIsNotObject = 2 ^ 11
    flIsSelfReferential = 2 ^ 12
    flIsNotSelfReferential = 2 ^ 13
    flIsString = 2 ^ 14
    flIsNotString = 2 ^ 15
    flIsVariant = 2 ^ 16
    flIsNotVariant = 2 ^ 17
End Enum

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
