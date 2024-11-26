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
Private pResult As Variant
Private pUtilities As IUtilities

Private Sub Class_Initialize()
    Set pUtilities = New cUtilities
End Sub

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
    Dim result As Variant
    
    result = pBeTestFuncs.GreaterThan(pNotBeValue, value, True)
    
    result = pUtilities.NegateIfBool(result)
    
    IBe_GreaterThan = result
End Function

Private Function IBe_LessThan(ByVal value As Double) As Variant
    Dim result As Variant
    result = pBeTestFuncs.LessThan(pNotBeValue, value, True)
    result = pUtilities.NegateIfBool(result)
    IBe_LessThan = result
End Function

Private Function IBe_EqualTo(ByVal value As Variant) As Variant
    Dim result As Variant
    
    result = pBeTestFuncs.EqualTo(pNotBeValue, value, True, updateFluentPath:=True)

    result = pUtilities.NegateIfBool(result)
    
    IBe_EqualTo = result
End Function

Private Function IBe_GreaterThanOrEqualTo(ByVal value As Double) As Variant
    Dim result As Variant
    result = pBeTestFuncs.GreaterThanOrEqualTo(pNotBeValue, value, True, updateFluentPath:=True)
    result = pUtilities.NegateIfBool(result)
    IBe_GreaterThanOrEqualTo = result
End Function

Private Function IBe_LessThanOrEqualTo(ByVal value As Double) As Variant
    Dim result As Variant
    result = pBeTestFuncs.LessThanOrEqualTo(pNotBeValue, value, True, updateFluentPath:=True)
    result = pUtilities.NegateIfBool(result)
    IBe_LessThanOrEqualTo = result
End Function

Private Function IBe_Something() As Variant
    Dim result As Variant
    result = pBeTestFuncs.Something(pNotBeValue, True)
    result = pUtilities.NegateIfBool(result)
    IBe_Something = result
End Function

Private Function IBe_Between(ByVal lowerVal As Double, ByVal higherVal As Double) As Variant
    Dim result As Variant
    result = pBeTestFuncs.Between(pNotBeValue, lowerVal, higherVal, True)
    result = pUtilities.NegateIfBool(result)
    IBe_Between = result
End Function

Private Function IBe_OneOf(ParamArray All() As Variant) As Variant
    Dim result As Variant
    result = pBeTestFuncs.OneOf(pNotBeValue, False, All)
    result = pUtilities.NegateIfBool(result)
    IBe_OneOf = result
End Function

Private Function IBe_InDataStructure(ByVal dataStructure As Variant) As Variant
    Dim result As Variant
    result = pBeTestFuncs.InDataStructure(pNotBeValue, dataStructure, True, updateFluentPath:=True)
    result = pUtilities.NegateIfBool(result)
    IBe_InDataStructure = result
End Function

Private Function IBe_InDataStructures(ParamArray DataStructures() As Variant) As Variant
    Dim result As Variant
    result = pBeTestFuncs.InDataStructures(pNotBeValue, True, DataStructures)
    result = pUtilities.NegateIfBool(result)
    IBe_InDataStructures = result
End Function

Private Function IBe_Alphabetic() As Variant
    Dim result As Variant
    result = pBeTestFuncs.Alphabetic(pNotBeValue, True)
    result = pUtilities.NegateIfBool(result)
    IBe_Alphabetic = result
End Function

Private Function IBe_Alphanumeric() As Variant
    Dim result As Variant
    result = pBeTestFuncs.Alphanumeric(pNotBeValue, True)
    result = pUtilities.NegateIfBool(result)
    IBe_Alphanumeric = result
End Function

Private Function IBe_Numeric() As Variant
    Dim result As Variant
    result = pBeTestFuncs.Numeric(pNotBeValue, True)
    result = pUtilities.NegateIfBool(result)
    IBe_Numeric = result
End Function

Private Function IBe_Erroneous() As Variant
    Dim result As Variant
    result = pBeTestFuncs.Erroneous(pNotBeValue, True)
    result = pUtilities.NegateIfBool(result)
    IBe_Erroneous = result
End Function

Private Function IBe_IdenticalTo(ByVal dataStructure As Variant) As Variant
    Dim result As Variant
    result = pBeTestFuncs.IdenticalTo(pNotBeValue, True, dataStructure, True)
    result = pUtilities.NegateIfBool(result)
    IBe_IdenticalTo = result
End Function

