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
Private pBeTestFuncs As IBeTestFuncs
Private pBeValue As Variant

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
    Set pBeTestFuncs = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pBeValue = value
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pBeValue = value
End Property

Private Function IBe_GreaterThan(ByVal value As Double) As Variant
    IBe_GreaterThan = pBeTestFuncs.GreaterThan(pBeValue, value)
End Function

Private Function IBe_LessThan(ByVal value As Double) As Variant
    IBe_LessThan = pBeTestFuncs.LessThan(pBeValue, value)
End Function

Private Function IBe_EqualTo(ByVal value As Variant) As Variant
    IBe_EqualTo = pBeTestFuncs.EqualTo(pBeValue, value, updateFluentPath:=True)
End Function

Private Function IBe_GreaterThanOrEqualTo(ByVal value As Double) As Variant
    IBe_GreaterThanOrEqualTo = pBeTestFuncs.GreaterThanOrEqualTo(pBeValue, value, updateFluentPath:=True)
End Function

Private Function IBe_LessThanOrEqualTo(ByVal value As Double) As Variant
    IBe_LessThanOrEqualTo = pBeTestFuncs.LessThanOrEqualTo(pBeValue, value, updateFluentPath:=True)
End Function

Private Function IBe_Something() As Variant
    IBe_Something = pBeTestFuncs.Something(pBeValue)
End Function

Private Function IBe_Between(ByVal lowerVal As Double, ByVal higherVal As Double) As Variant
    IBe_Between = pBeTestFuncs.Between(pBeValue, lowerVal, higherVal)
End Function

Private Function IBe_OneOf(ParamArray All() As Variant) As Variant
    IBe_OneOf = pBeTestFuncs.OneOf(pBeValue, False, All)
End Function

Private Function IBe_InDataStructure(ByVal dataStructure As Variant) As Variant
    IBe_InDataStructure = pBeTestFuncs.InDataStructure(pBeValue, dataStructure, updateFluentPath:=True)
End Function

Private Function IBe_InDataStructures(ParamArray DataStructures() As Variant) As Variant
    IBe_InDataStructures = pBeTestFuncs.InDataStructures(pBeValue, False, DataStructures)
End Function

Private Function IBe_Alphabetic() As Variant
    IBe_Alphabetic = pBeTestFuncs.Alphabetic(pBeValue)
End Function

Private Function IBe_Alphanumeric() As Variant
    IBe_Alphanumeric = pBeTestFuncs.Alphanumeric(pBeValue)
End Function

Private Function IBe_Numeric() As Variant
    IBe_Numeric = pBeTestFuncs.Numeric(pBeValue)
End Function

Private Function IBe_Erroneous() As Variant
    IBe_Erroneous = pBeTestFuncs.Erroneous(pBeValue)
End Function

Private Function IBe_IdenticalTo(ByVal dataStructure As Variant) As Variant
    IBe_IdenticalTo = pBeTestFuncs.IdenticalTo(pBeValue, False, dataStructure, True)
End Function

Private Sub Class_Terminate()
    Set pTestingFunctions = Nothing
    Set pBeTestFuncs = Nothing
End Sub
