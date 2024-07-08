VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNotHave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHave
Implements ISetTestFuncsAndValue

Private pNotHaveVal As Variant
Private pHaveTestFuncs As IHaveTestFuncs
Private pTestingFunctions As ITestingFunctions
'Private pNegate As INegate
Private pUtilities As IUtilities

Private Sub Class_Initialize()
'    Set pNegate = New cNegate
    Set pUtilities = New cUtilities
End Sub

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
    Set pHaveTestFuncs = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pNotHaveVal = pTestingFunctions.TestValue
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pNotHaveVal = pTestingFunctions.TestValue
End Property

Private Function IHave_LengthOf(ByVal value As Long) As Variant
    Dim Result As Variant
    Result = pHaveTestFuncs.LengthOf(pNotHaveVal, value, True)
    Result = pUtilities.NegateIfBool(Result)
    IHave_LengthOf = Result
End Function

Private Function IHave_MaxLengthOf(ByVal value As Long) As Variant
    Dim Result As Variant
    Result = pHaveTestFuncs.MaxLengthOf(pNotHaveVal, value, True, updateFluentPath:=True)
    Result = pUtilities.NegateIfBool(Result)
    IHave_MaxLengthOf = Result
End Function

Private Function IHave_MinLengthOf(ByVal value As Long) As Variant
    Dim Result As Variant
    Result = pHaveTestFuncs.MinLengthOf(pNotHaveVal, value, True, updateFluentPath:=True)
    Result = pUtilities.NegateIfBool(Result)
    IHave_MinLengthOf = Result
End Function

Private Function IHave_LengthBetween(ByVal lowVal As Long, ByVal highVal As Long) As Variant
    Dim Result As Variant
    Result = pHaveTestFuncs.LengthBetween(pNotHaveVal, lowVal, highVal, True)
    Result = pUtilities.NegateIfBool(Result)
    IHave_LengthBetween = Result
End Function

Private Function IHave_ErrorNumberOf(ByVal value As Long) As Variant
    Dim Result As Variant
    Result = pHaveTestFuncs.ErrorNumberOf(pNotHaveVal, value, True)
    Result = pUtilities.NegateIfBool(Result)
    IHave_ErrorNumberOf = Result
End Function

Private Function IHave_ErrorDescriptionOf(ByVal value As String) As Variant
    Dim Result As Variant
    Result = pHaveTestFuncs.ErrorDescriptionOf(pNotHaveVal, value, True)
    Result = pUtilities.NegateIfBool(Result)
    IHave_ErrorDescriptionOf = Result
End Function

Private Function IHave_SameTypeAs(ByVal All As Variant) As Variant
    Dim Result As Variant
    Result = pHaveTestFuncs.SameTypeAs(pNotHaveVal, All, True)
    Result = pUtilities.NegateIfBool(Result)
    IHave_SameTypeAs = Result
End Function

Private Function IHave_ExactSameElementsAs(ByVal dataStructure As Variant) As Variant
    Dim Result As Variant
    Result = pHaveTestFuncs.ExactSameElementsAs(pNotHaveVal, True, True, dataStructure)
    Result = pUtilities.NegateIfBool(Result)
    IHave_ExactSameElementsAs = Result
End Function

Private Function IHave_SameUniqueElementsAs(ByVal dataStructure As Variant) As Variant
    Dim Result As Variant
    Result = pHaveTestFuncs.SameUniqueElementsAs(pNotHaveVal, True, dataStructure)
    Result = pUtilities.NegateIfBool(Result)
    IHave_SameUniqueElementsAs = Result
End Function

Public Function IHave_SameElementsAs(ByVal dataStructure As Variant) As Variant
    Dim Result As Variant
    Result = pHaveTestFuncs.SameElementsAs(pNotHaveVal, True, dataStructure)
    Result = pUtilities.NegateIfBool(Result)
    IHave_SameElementsAs = Result
End Function
