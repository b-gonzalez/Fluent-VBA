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
Private pUtilities As IUtilities

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
    Dim result As Variant
    result = pHaveTestFuncs.LengthOf(pNotHaveVal, value, True)
    result = pUtilities.NegateIfBool(result)
    IHave_LengthOf = result
End Function

Private Function IHave_MaxLengthOf(ByVal value As Long) As Variant
    Dim result As Variant
    result = pHaveTestFuncs.MaxLengthOf(pNotHaveVal, value, True, updateFluentPath:=True)
    result = pUtilities.NegateIfBool(result)
    IHave_MaxLengthOf = result
End Function

Private Function IHave_MinLengthOf(ByVal value As Long) As Variant
    Dim result As Variant
    result = pHaveTestFuncs.MinLengthOf(pNotHaveVal, value, True, updateFluentPath:=True)
    result = pUtilities.NegateIfBool(result)
    IHave_MinLengthOf = result
End Function

Private Function IHave_LengthBetween(ByVal lowVal As Long, ByVal highVal As Long) As Variant
    Dim result As Variant
    result = pHaveTestFuncs.LengthBetween(pNotHaveVal, lowVal, highVal, True)
    result = pUtilities.NegateIfBool(result)
    IHave_LengthBetween = result
End Function

Private Function IHave_ErrorNumberOf(ByVal value As Long) As Variant
    Dim result As Variant
    result = pHaveTestFuncs.ErrorNumberOf(pNotHaveVal, value, True)
    result = pUtilities.NegateIfBool(result)
    IHave_ErrorNumberOf = result
End Function

Private Function IHave_ErrorDescriptionOf(ByVal value As String) As Variant
    Dim result As Variant
    result = pHaveTestFuncs.ErrorDescriptionOf(pNotHaveVal, value, True)
    result = pUtilities.NegateIfBool(result)
    IHave_ErrorDescriptionOf = result
End Function

Private Function IHave_SameTypeAs(ByVal All As Variant) As Variant
    Dim result As Variant
    result = pHaveTestFuncs.SameTypeAs(pNotHaveVal, All, True)
    result = pUtilities.NegateIfBool(result)
    IHave_SameTypeAs = result
End Function

Private Function IHave_ExactSameElementsAs(ByVal dataStructure As Variant) As Variant
    Dim result As Variant
    result = pHaveTestFuncs.ExactSameElementsAs(pNotHaveVal, True, True, dataStructure)
    result = pUtilities.NegateIfBool(result)
    IHave_ExactSameElementsAs = result
End Function

Private Function IHave_SameUniqueElementsAs(ByVal dataStructure As Variant) As Variant
    Dim result As Variant
    result = pHaveTestFuncs.SameUniqueElementsAs(pNotHaveVal, True, dataStructure)
    result = pUtilities.NegateIfBool(result)
    IHave_SameUniqueElementsAs = result
End Function

Private Function IHave_SameElementsAs(ByVal dataStructure As Variant) As Variant
    Dim result As Variant
    result = pHaveTestFuncs.SameElementsAs(pNotHaveVal, True, dataStructure)
    result = pUtilities.NegateIfBool(result)
    IHave_SameElementsAs = result
End Function

Private Function IHave_Procedure( _
    ByVal procName As String, _
    ByVal callType As VBA.VbCallType _
) As Variant
    Dim result As Variant
    
    result = pHaveTestFuncs.Procedure(pNotHaveVal, procName, callType, True)
    result = pUtilities.NegateIfBool(result)
    IHave_Procedure = result
End Function

Private Function IHave_Elements(ParamArray args() As Variant) As Variant
    Dim tempArgs() As Variant
    Dim result As Variant
    
    tempArgs = args
    
    result = pHaveTestFuncs.Elements(pNotHaveVal, tempArgs, True, True)
    result = pUtilities.NegateIfBool(result)
    IHave_Elements = result
End Function

Private Function IHave_ElementsInDataStructure(ByVal dataStructure As Variant) As Variant
    Dim result As Variant
        
    result = pHaveTestFuncs.ElementsInDataStructure(pNotHaveVal, dataStructure, True)
    result = pUtilities.NegateIfBool(result)
    IHave_ElementsInDataStructure = result
End Function

Private Function IHave_DepthCountOf( _
    ByVal value As Long _
) As Variant
    Dim result As Variant
    
    result = pHaveTestFuncs.DepthCountOf(pNotHaveVal, value, True)
    result = pUtilities.NegateIfBool(result)
    IHave_DepthCountOf = result
End Function

Private Function IHave_NestedCountOf( _
    ByVal value As Long _
) As Variant
    Dim result As Variant
    
    result = pHaveTestFuncs.NestedCountOf(pNotHaveVal, value, True)
    result = pUtilities.NegateIfBool(result)
    IHave_NestedCountOf = result
End Function

Private Sub Class_Initialize()
    Set pUtilities = New cUtilities
End Sub

Private Sub Class_Terminate()
    Set pUtilities = Nothing
    Set pTestingFunctions = Nothing
    Set pHaveTestFuncs = Nothing
End Sub
