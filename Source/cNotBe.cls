VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNotBe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IBe
Implements ISetTestFuncsAndValue

Private pNotBeValue As Variant
Private pBe As IBe
Private pTestingFunctions As ITestingFunctions

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pNotBeValue = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pNotBeValue = value
End Property

Private Function IBe_GreaterThan(ByVal value As Double) As Variant
    IBe_GreaterThan = Not pTestingFunctions.GreaterThan(pNotBeValue, value, True)
End Function

Private Function IBe_LessThan(ByVal value As Double) As Variant
    IBe_LessThan = Not pTestingFunctions.LessThan(pNotBeValue, value, True)
End Function

Private Function IBe_EqualTo(ByVal value As Variant) As Variant
    IBe_EqualTo = Not pTestingFunctions.EqualTo(pNotBeValue, value, True, updateFluentPath:=True)
End Function

Private Function IBe_GreaterThanOrEqualTo(ByVal value As Double) As Variant
    IBe_GreaterThanOrEqualTo = Not pTestingFunctions.GreaterThanOrEqualTo(pNotBeValue, value, True, updateFluentPath:=True)
End Function

Private Function IBe_LessThanOrEqualTo(ByVal value As Double) As Variant
    IBe_LessThanOrEqualTo = Not pTestingFunctions.LessThanOrEqualTo(pNotBeValue, value, True, updateFluentPath:=True)
End Function

Private Function IBe_Something() As Variant
    IBe_Something = Not pTestingFunctions.Something(pNotBeValue, True)
End Function

Private Function IBe_Between(ByVal LowerVal As Double, ByVal HigherVal As Double) As Variant
    IBe_Between = Not pTestingFunctions.Between(CDbl(pNotBeValue), LowerVal, HigherVal, True)
End Function

Private Function IBe_OneOf(ParamArray Vals() As Variant) As Variant
    IBe_OneOf = Not pTestingFunctions.OneOf(pNotBeValue, False, Vals)
End Function

Private Function IBe_InDataStructure(ByVal DataStructure As Variant, Optional ByVal method As flAlgorithm = flRecursive) As Variant
    IBe_InDataStructure = Not pTestingFunctions.InDataStructure(pNotBeValue, DataStructure, method, True, updateFluentPath:=True)
End Function

Private Function IBe_InDataStructures(ByVal method As flAlgorithm, ParamArray DataStructures() As Variant) As Variant
    IBe_InDataStructures = Not pTestingFunctions.InDataStructures(pNotBeValue, method, True, DataStructures)
End Function

Private Function IBe_Alphabetic() As Variant
    IBe_Alphabetic = Not pTestingFunctions.Alphabetic(pNotBeValue, True)
End Function

Private Function IBe_Alphanumeric() As Variant
    IBe_Alphanumeric = Not pTestingFunctions.Alphanumeric(pNotBeValue, True)
End Function

Private Function IBe_Numeric() As Variant
    IBe_Numeric = Not pTestingFunctions.Numeric(pNotBeValue, True)
End Function

Private Function IBe_Erroneous() As Variant
    IBe_Erroneous = Not pTestingFunctions.Erroneous(pNotBeValue, True)
End Function

Private Function IBe_IdenticalTo(ByVal value As Variant) As Variant
'    Dim temp As Variant
'    temp = args
    IBe_IdenticalTo = Not pTestingFunctions.IdenticalTo(pNotBeValue, True, value, True)
End Function
