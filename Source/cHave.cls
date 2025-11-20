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
    pHaveVal = pTestingFunctions.testValue
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pHaveVal = pTestingFunctions.testValue
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
    IHave_ExactSameElementsAs = pHaveTestFuncs.ExactSameElementsAs(pHaveVal, False, True, dataStructure)
End Function

Private Function IHave_SameUniqueElementsAs(ByVal dataStructure As Variant) As Variant
    IHave_SameUniqueElementsAs = pHaveTestFuncs.SameUniqueElementsAs(pHaveVal, False, dataStructure)
End Function

Private Function IHave_SameElementsAs(ByVal dataStructure As Variant) As Variant
    IHave_SameElementsAs = pHaveTestFuncs.SameElementsAs(pHaveVal, False, dataStructure)
End Function

Private Function IHave_Procedure( _
    ByVal procName As String, _
    ByVal callType As VBA.VbCallType _
) As Variant
    IHave_Procedure = pHaveTestFuncs.Procedure(pHaveVal, procName, callType, False)
End Function

Private Function IHave_Elements(ParamArray args() As Variant) As Variant
    Dim tempArgs() As Variant
    
    tempArgs = args
    
    IHave_Elements = pHaveTestFuncs.Elements(pHaveVal, tempArgs, False, True)
End Function

Private Function IHave_ElementsInDataStructure(ByVal dataStructure As Variant) As Variant
    IHave_ElementsInDataStructure = pHaveTestFuncs.ElementsInDataStructure(pHaveVal, dataStructure, False)
End Function

Private Function IHave_DepthCountOf( _
    ByVal value As Long _
) As Variant
    IHave_DepthCountOf = pHaveTestFuncs.DepthCountOf(pHaveVal, value, False)
End Function

Private Function IHave_NestedCountOf( _
    ByVal value As Long _
) As Variant
    IHave_NestedCountOf = pHaveTestFuncs.NestedCountOf(pHaveVal, value, False)
End Function

Private Sub Class_Terminate()
    Set pTestingFunctions = Nothing
    Set pHaveTestFuncs = Nothing
End Sub
