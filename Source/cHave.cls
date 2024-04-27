VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHave
Implements ISetTestFuncsAndValue

Private pTestingFunctions As ITestingFunctions
Private pHaveVal As Variant

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pHaveVal = pTestingFunctions.TestValue
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pHaveVal = pTestingFunctions.TestValue
End Property

Private Function IHave_LengthOf(ByVal value As Long) As Variant
    IHave_LengthOf = pTestingFunctions.LengthOf(pHaveVal, value)
End Function

Private Function IHave_MaxLengthOf(ByVal value As Long) As Variant
    IHave_MaxLengthOf = pTestingFunctions.MaxLengthOf(pHaveVal, value, updateFluentPath:=True)
End Function

Private Function IHave_MinLengthOf(ByVal value As Long) As Variant
    IHave_MinLengthOf = pTestingFunctions.MinLengthOf(pHaveVal, value, updateFluentPath:=True)
End Function

Private Function IHave_LengthBetween(ByVal lowVal As Long, ByVal highVal As Long) As Variant
    IHave_LengthBetween = pTestingFunctions.LengthBetween(pHaveVal, lowVal, highVal)
End Function

Private Function IHave_ErrorNumberOf(ByVal value As Long) As Variant
    IHave_ErrorNumberOf = pTestingFunctions.ErrorNumberOf(pHaveVal, value)
End Function

Private Function IHave_ErrorDescriptionOf(ByVal value As String) As Variant
    IHave_ErrorDescriptionOf = pTestingFunctions.ErrorDescriptionOf(pHaveVal, value)
End Function

Private Function IHave_SameTypeAs(ByVal value As Variant) As Variant
    IHave_SameTypeAs = pTestingFunctions.SameTypeAs(pHaveVal, value)
End Function

Private Function IHave_ExactSameElementsAs(ByVal value As Variant) As Variant
'    Dim temp As Variant
'    temp = args
    IHave_ExactSameElementsAs = pTestingFunctions.ExactSameElementsAs(pHaveVal, False, True, value)
End Function

Private Function IHave_SameUniqueElementsAs(ByVal value As Variant) As Variant
    IHave_SameUniqueElementsAs = pTestingFunctions.SameUniqueElementsAs(pHaveVal, False, value)
End Function

Public Function IHave_SameElementsAs(ByVal value As Variant) As Variant
    IHave_SameElementsAs = pTestingFunctions.SameElementsAs(pHaveVal, False, value)
End Function
