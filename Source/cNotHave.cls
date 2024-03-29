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
Implements ISetExpressionAndValue

Private pNotHaveVal As Variant
Private pExpressions As cExpressions

Private Property Set ISetExpressionAndValue_setExpr(ByVal value As cExpressions)
    Set pExpressions = value
End Property

Private Property Let ISetExpressionAndValue_testValue(ByVal value As Variant)
    pNotHaveVal = pExpressions.TestValue
End Property

Private Property Set ISetExpressionAndValue_testValue(ByVal value As Variant)
    Set pNotHaveVal = pExpressions.TestValue
End Property

Private Function IHave_LengthOf(ByVal value As Double) As Boolean
    IHave_LengthOf = Not pExpressions.LengthOf(CDbl(pNotHaveVal), value, True)
End Function

Private Function IHave_MaxLengthOf(ByVal value As Double) As Boolean
    IHave_MaxLengthOf = Not pExpressions.MaxLengthOf(CDbl(pNotHaveVal), value, True, updateFluentPath:=True)
End Function

Private Function IHave_MinLengthOf(ByVal value As Double) As Boolean
    IHave_MinLengthOf = Not pExpressions.MinLengthOf(CDbl(pNotHaveVal), value, True, updateFluentPath:=True)
End Function

Private Function IHave_LengthBetween(ByVal lowVal As Double, ByVal highVal As Double) As Boolean
    IHave_LengthBetween = Not pExpressions.LengthBetween(CDbl(pNotHaveVal), lowVal, highVal, True)
End Function

Private Function IHave_ErrorNumberOf(ByVal value As Long) As Boolean
    IHave_ErrorNumberOf = Not pExpressions.ErrorNumberOf(pNotHaveVal, value, True)
End Function

Private Function IHave_ErrorDescriptionOf(ByVal value As String) As Boolean
    IHave_ErrorDescriptionOf = Not pExpressions.ErrorDescriptionOf(pNotHaveVal, value, True)
End Function

Private Function IHave_SameTypeAs(ByVal value As Variant) As Boolean
    IHave_SameTypeAs = Not pExpressions.SameTypeAs(pNotHaveVal, value, True)
End Function

Private Function IHave_ExactSameElementsAs(ByVal value As Variant) As Boolean
'    Dim temp As Variant
'    temp = args
    IHave_ExactSameElementsAs = Not pExpressions.ExactSameElementsAs(pNotHaveVal, True, True, value)
End Function

Private Function IHave_SameUniqueElementsAs(ByVal value As Variant) As Boolean
    IHave_SameUniqueElementsAs = Not pExpressions.SameUniqueElementsAs(pNotHaveVal, True, value)
End Function

Public Function IHave_SameElementsAs(ByVal value As Variant) As Boolean
    IHave_SameElementsAs = Not pExpressions.SameElementsAs(pNotHaveVal, True, value)
End Function
