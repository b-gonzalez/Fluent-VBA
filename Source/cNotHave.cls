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
    IHave_LengthOf = Not pHaveTestFuncs.LengthOf(pNotHaveVal, value, True)
End Function

Private Function IHave_MaxLengthOf(ByVal value As Long) As Variant
    IHave_MaxLengthOf = Not pHaveTestFuncs.MaxLengthOf(pNotHaveVal, value, True, updateFluentPath:=True)
End Function

Private Function IHave_MinLengthOf(ByVal value As Long) As Variant
    IHave_MinLengthOf = Not pHaveTestFuncs.MinLengthOf(pNotHaveVal, value, True, updateFluentPath:=True)
End Function

Private Function IHave_LengthBetween(ByVal lowVal As Long, ByVal highVal As Long) As Variant
    IHave_LengthBetween = Not pHaveTestFuncs.LengthBetween(pNotHaveVal, lowVal, highVal, True)
End Function

Private Function IHave_ErrorNumberOf(ByVal value As Long) As Variant
    IHave_ErrorNumberOf = Not pHaveTestFuncs.ErrorNumberOf(pNotHaveVal, value, True)
End Function

Private Function IHave_ErrorDescriptionOf(ByVal value As String) As Variant
    IHave_ErrorDescriptionOf = Not pHaveTestFuncs.ErrorDescriptionOf(pNotHaveVal, value, True)
End Function

Private Function IHave_SameTypeAs(ByVal All As Variant) As Variant
    IHave_SameTypeAs = Not pHaveTestFuncs.SameTypeAs(pNotHaveVal, All, True)
End Function

Private Function IHave_ExactSameElementsAs(ByVal dataStructure As Variant) As Variant
'    Dim temp As Variant
'    temp = args
    IHave_ExactSameElementsAs = Not pHaveTestFuncs.ExactSameElementsAs(pNotHaveVal, True, True, dataStructure)
End Function

Private Function IHave_SameUniqueElementsAs(ByVal dataStructure As Variant) As Variant
    IHave_SameUniqueElementsAs = Not pHaveTestFuncs.SameUniqueElementsAs(pNotHaveVal, True, dataStructure)
End Function

Public Function IHave_SameElementsAs(ByVal dataStructure As Variant) As Variant
    IHave_SameElementsAs = Not pHaveTestFuncs.SameElementsAs(pNotHaveVal, True, dataStructure)
End Function
