VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IBe
Implements ISetTestFuncsAndValue

Private pTestingFunctions As ITestingFunctions
Private pBeValue As Variant

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pBeValue = value
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pBeValue = value
End Property

Private Function IBe_GreaterThan(ByVal value As Double) As Variant
    IBe_GreaterThan = pTestingFunctions.GreaterThan(pBeValue, value)
End Function

Private Function IBe_LessThan(ByVal value As Double) As Variant
    IBe_LessThan = pTestingFunctions.LessThan(pBeValue, value)
End Function

Private Function IBe_EqualTo(ByVal value As Variant) As Variant
    IBe_EqualTo = pTestingFunctions.EqualTo(pBeValue, value, updateFluentPath:=True)
End Function

Private Function IBe_GreaterThanOrEqualTo(ByVal value As Double) As Variant
    IBe_GreaterThanOrEqualTo = pTestingFunctions.GreaterThanOrEqualTo(pBeValue, value, updateFluentPath:=True)
End Function

Private Function IBe_LessThanOrEqualTo(ByVal value As Double) As Variant
    IBe_LessThanOrEqualTo = pTestingFunctions.LessThanOrEqualTo(pBeValue, value, updateFluentPath:=True)
End Function

Private Function IBe_Something() As Variant
    IBe_Something = pTestingFunctions.Something(pBeValue)
End Function

Private Function IBe_Between(ByVal LowerVal As Double, ByVal HigherVal As Double) As Variant
    IBe_Between = pTestingFunctions.Between(pBeValue, LowerVal, HigherVal)
End Function

Private Function IBe_OneOf(ParamArray Vals() As Variant) As Variant
    IBe_OneOf = pTestingFunctions.OneOf(pBeValue, False, Vals)
End Function

Private Function IBe_InDataStructure(ByVal DataStructure As Variant, Optional ByVal method As flAlgorithm = flRecursive) As Variant
    IBe_InDataStructure = pTestingFunctions.InDataStructure(pBeValue, DataStructure, method, updateFluentPath:=True)
End Function

Private Function IBe_InDataStructures(ByVal method As flAlgorithm, ParamArray DataStructures() As Variant) As Variant
    IBe_InDataStructures = pTestingFunctions.InDataStructures(pBeValue, method, False, DataStructures)
End Function

Private Function IBe_Alphabetic() As Variant
    IBe_Alphabetic = pTestingFunctions.Alphabetic(pBeValue)
End Function

Private Function IBe_Alphanumeric() As Variant
    IBe_Alphanumeric = pTestingFunctions.Alphanumeric(pBeValue)
End Function

Private Function IBe_Numeric() As Variant
    IBe_Numeric = pTestingFunctions.Numeric(pBeValue)
End Function

Private Function IBe_Erroneous() As Variant
    IBe_Erroneous = pTestingFunctions.Erroneous(pBeValue)
End Function

Private Function IBe_IdenticalTo(ByVal value As Variant) As Variant
'    Dim temp As Variant
'    temp = args
    IBe_IdenticalTo = pTestingFunctions.IdenticalTo(pBeValue, False, value, True)
End Function
