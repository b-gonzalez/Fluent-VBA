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
Private pHaveTestFuncs As IHaveTestFuncs
Private pHaveVal As Variant

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
    Set pHaveTestFuncs = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pHaveVal = pTestingFunctions.TestValue
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pHaveVal = pTestingFunctions.TestValue
End Property

Private Function IHave_LengthOf(ByVal value As Long) As Variant
    IHave_LengthOf = pHaveTestFuncs.LengthOf(pHaveVal, value)
End Function

Private Function IHave_MaxLengthOf(ByVal value As Long) As Variant
    IHave_MaxLengthOf = pHaveTestFuncs.MaxLengthOf(pHaveVal, value, updateFluentPath:=True)
End Function

Private Function IHave_MinLengthOf(ByVal value As Long) As Variant
    IHave_MinLengthOf = pHaveTestFuncs.MinLengthOf(pHaveVal, value, updateFluentPath:=True)
End Function

Private Function IHave_LengthBetween(ByVal lowVal As Long, ByVal highVal As Long) As Variant
    IHave_LengthBetween = pHaveTestFuncs.LengthBetween(pHaveVal, lowVal, highVal)
End Function

Private Function IHave_ErrorNumberOf(ByVal value As Long) As Variant
    IHave_ErrorNumberOf = pHaveTestFuncs.ErrorNumberOf(pHaveVal, value)
End Function

Private Function IHave_ErrorDescriptionOf(ByVal value As String) As Variant
    IHave_ErrorDescriptionOf = pHaveTestFuncs.ErrorDescriptionOf(pHaveVal, value)
End Function

Private Function IHave_SameTypeAs(ByVal All As Variant) As Variant
    IHave_SameTypeAs = pHaveTestFuncs.SameTypeAs(pHaveVal, All)
End Function

Private Function IHave_ExactSameElementsAs(ByVal dataStructure As Variant) As Variant
'    Dim temp As Variant
'    temp = args
    IHave_ExactSameElementsAs = pHaveTestFuncs.ExactSameElementsAs(pHaveVal, False, True, dataStructure)
End Function

Private Function IHave_SameUniqueElementsAs(ByVal dataStructure As Variant) As Variant
    IHave_SameUniqueElementsAs = pHaveTestFuncs.SameUniqueElementsAs(pHaveVal, False, dataStructure)
End Function

Private Function IHave_SameElementsAs(ByVal dataStructure As Variant) As Variant
    IHave_SameElementsAs = pHaveTestFuncs.SameElementsAs(pHaveVal, False, dataStructure)
End Function

Private Function IHave_Procedure( _
    procName As String, _
    callType As VBA.VbCallType _
) As Variant
    IHave_Procedure = pHaveTestFuncs.Procedure(pHaveVal, procName, callType, False)
End Function

Private Function IHave_Elements(ParamArray args() As Variant)
    Dim tempArgs() As Variant
    
    tempArgs = args
    
    IHave_Elements = pHaveTestFuncs.Elements(pHaveVal, tempArgs, False)
End Function
