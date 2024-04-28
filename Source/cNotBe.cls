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

Private pTestingFunctions As ITestingFunctions
Private pBeTestFuncs As IBeTestFuncs

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
    Set pBeTestFuncs = value
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pNotBeValue = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pNotBeValue = value
End Property

Private Function IBe_GreaterThan(ByVal value As Double) As Variant
    IBe_GreaterThan = Not pBeTestFuncs.GreaterThan(pNotBeValue, value, True)
End Function

Private Function IBe_LessThan(ByVal value As Double) As Variant
    IBe_LessThan = Not pBeTestFuncs.LessThan(pNotBeValue, value, True)
End Function

Private Function IBe_EqualTo(ByVal value As Variant) As Variant
    IBe_EqualTo = Not pBeTestFuncs.EqualTo(pNotBeValue, value, True, updateFluentPath:=True)
End Function

Private Function IBe_GreaterThanOrEqualTo(ByVal value As Double) As Variant
    IBe_GreaterThanOrEqualTo = Not pBeTestFuncs.GreaterThanOrEqualTo(pNotBeValue, value, True, updateFluentPath:=True)
End Function

Private Function IBe_LessThanOrEqualTo(ByVal value As Double) As Variant
    IBe_LessThanOrEqualTo = Not pBeTestFuncs.LessThanOrEqualTo(pNotBeValue, value, True, updateFluentPath:=True)
End Function

Private Function IBe_Something() As Variant
    IBe_Something = Not pBeTestFuncs.Something(pNotBeValue, True)
End Function

Private Function IBe_Between(ByVal LowerVal As Double, ByVal HigherVal As Double) As Variant
    IBe_Between = Not pBeTestFuncs.Between(CDbl(pNotBeValue), LowerVal, HigherVal, True)
End Function

Private Function IBe_OneOf(ParamArray All() As Variant) As Variant
    IBe_OneOf = Not pBeTestFuncs.OneOf(pNotBeValue, False, All)
End Function

Private Function IBe_InDataStructure(ByVal DataStructure As Variant, Optional ByVal method As flAlgorithm = flRecursive) As Variant
    IBe_InDataStructure = Not pBeTestFuncs.InDataStructure(pNotBeValue, DataStructure, method, True, updateFluentPath:=True)
End Function

Private Function IBe_InDataStructures(ByVal method As flAlgorithm, ParamArray DataStructures() As Variant) As Variant
    IBe_InDataStructures = Not pBeTestFuncs.InDataStructures(pNotBeValue, method, True, DataStructures)
End Function

Private Function IBe_Alphabetic() As Variant
    IBe_Alphabetic = Not pBeTestFuncs.Alphabetic(pNotBeValue, True)
End Function

Private Function IBe_Alphanumeric() As Variant
    IBe_Alphanumeric = Not pBeTestFuncs.Alphanumeric(pNotBeValue, True)
End Function

Private Function IBe_Numeric() As Variant
    IBe_Numeric = Not pBeTestFuncs.Numeric(pNotBeValue, True)
End Function

Private Function IBe_Erroneous() As Variant
    IBe_Erroneous = Not pBeTestFuncs.Erroneous(pNotBeValue, True)
End Function

Private Function IBe_IdenticalTo(ByVal DataStructure As Variant) As Variant
'    Dim temp As Variant
'    temp = args
    IBe_IdenticalTo = Not pBeTestFuncs.IdenticalTo(pNotBeValue, True, DataStructure, True)
End Function
