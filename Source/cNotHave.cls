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
Private pTestingFunctions As cTestingFunctions

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pNotHaveVal = pTestingFunctions.TestValue
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pNotHaveVal = pTestingFunctions.TestValue
End Property

Private Function IHave_LengthOf(ByVal value As Long) As Variant
    IHave_LengthOf = Not pTestingFunctions.LengthOf(pNotHaveVal, value, True)
End Function

Private Function IHave_MaxLengthOf(ByVal value As Long) As Variant
    IHave_MaxLengthOf = Not pTestingFunctions.MaxLengthOf(pNotHaveVal, value, True, updateFluentPath:=True)
End Function

Private Function IHave_MinLengthOf(ByVal value As Long) As Variant
    IHave_MinLengthOf = Not pTestingFunctions.MinLengthOf(pNotHaveVal, value, True, updateFluentPath:=True)
End Function

Private Function IHave_LengthBetween(ByVal lowVal As Long, ByVal highVal As Long) As Variant
    IHave_LengthBetween = Not pTestingFunctions.LengthBetween(pNotHaveVal, lowVal, highVal, True)
End Function

Private Function IHave_ErrorNumberOf(ByVal value As Long) As Variant
    IHave_ErrorNumberOf = Not pTestingFunctions.ErrorNumberOf(pNotHaveVal, value, True)
End Function

Private Function IHave_ErrorDescriptionOf(ByVal value As String) As Variant
    IHave_ErrorDescriptionOf = Not pTestingFunctions.ErrorDescriptionOf(pNotHaveVal, value, True)
End Function

Private Function IHave_SameTypeAs(ByVal value As Variant) As Variant
    IHave_SameTypeAs = Not pTestingFunctions.SameTypeAs(pNotHaveVal, value, True)
End Function

Private Function IHave_ExactSameElementsAs(ByVal value As Variant) As Variant
'    Dim temp As Variant
'    temp = args
    IHave_ExactSameElementsAs = Not pTestingFunctions.ExactSameElementsAs(pNotHaveVal, True, True, value)
End Function

Private Function IHave_SameUniqueElementsAs(ByVal value As Variant) As Variant
    IHave_SameUniqueElementsAs = Not pTestingFunctions.SameUniqueElementsAs(pNotHaveVal, True, value)
End Function

Public Function IHave_SameElementsAs(ByVal value As Variant) As Variant
    IHave_SameElementsAs = Not pTestingFunctions.SameElementsAs(pNotHaveVal, True, value)
End Function
