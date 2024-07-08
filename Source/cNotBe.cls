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
    Dim Result As Variant
    
    Result = pBeTestFuncs.GreaterThan(pNotBeValue, value, True)
    
    Result = pUtilities.NegateIfBool(Result)
    
    IBe_GreaterThan = Result
End Function

Private Function IBe_LessThan(ByVal value As Double) As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.LessThan(pNotBeValue, value, True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_LessThan = Result
End Function

Private Function IBe_EqualTo(ByVal value As Variant) As Variant
    Dim Result As Variant
    
    Result = pBeTestFuncs.EqualTo(pNotBeValue, value, True, updateFluentPath:=True)

    Result = pUtilities.NegateIfBool(Result)
    
    IBe_EqualTo = Result
End Function

Private Function IBe_GreaterThanOrEqualTo(ByVal value As Double) As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.GreaterThanOrEqualTo(pNotBeValue, value, True, updateFluentPath:=True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_GreaterThanOrEqualTo = Result
End Function

Private Function IBe_LessThanOrEqualTo(ByVal value As Double) As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.LessThanOrEqualTo(pNotBeValue, value, True, updateFluentPath:=True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_LessThanOrEqualTo = Result
End Function

Private Function IBe_Something() As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.Something(pNotBeValue, True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_Something = Result
End Function

Private Function IBe_Between(ByVal LowerVal As Double, ByVal HigherVal As Double) As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.Between(pNotBeValue, LowerVal, HigherVal, True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_Between = Result
End Function

Private Function IBe_OneOf(ParamArray All() As Variant) As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.OneOf(pNotBeValue, False, All)
    Result = pUtilities.NegateIfBool(Result)
    IBe_OneOf = Result
End Function

Private Function IBe_InDataStructure(ByVal dataStructure As Variant) As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.InDataStructure(pNotBeValue, dataStructure, True, updateFluentPath:=True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_InDataStructure = Result
End Function

Private Function IBe_InDataStructures(ParamArray DataStructures() As Variant) As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.InDataStructures(pNotBeValue, True, DataStructures)
    Result = pUtilities.NegateIfBool(Result)
    IBe_InDataStructures = Result
End Function

Private Function IBe_Alphabetic() As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.Alphabetic(pNotBeValue, True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_Alphabetic = Result
End Function

Private Function IBe_Alphanumeric() As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.Alphanumeric(pNotBeValue, True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_Alphanumeric = Result
End Function

Private Function IBe_Numeric() As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.Numeric(pNotBeValue, True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_Numeric = Result
End Function

Private Function IBe_Erroneous() As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.Erroneous(pNotBeValue, True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_Erroneous = Result
End Function

Private Function IBe_IdenticalTo(ByVal dataStructure As Variant) As Variant
    Dim Result As Variant
    Result = pBeTestFuncs.IdenticalTo(pNotBeValue, True, dataStructure, True)
    Result = pUtilities.NegateIfBool(Result)
    IBe_IdenticalTo = Result
End Function

