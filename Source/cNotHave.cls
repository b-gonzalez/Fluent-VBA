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

Private Property Set ISetExpressionAndValue_setExpr(Value As cExpressions)
    Set pExpressions = Value
End Property

Private Property Let ISetExpressionAndValue_testValue(Value As Variant)
    pNotHaveVal = pExpressions.TestValue
End Property

Private Property Set ISetExpressionAndValue_testValue(Value As Variant)
    'Intentionally left blank
End Property

Private Function IHave_LengthOf(Value As Double) As Boolean
    IHave_LengthOf = Not pExpressions.LengthOf(CDbl(pNotHaveVal), Value, True)
End Function

Private Function IHave_MaxLengthOf(Value As Double) As Boolean
    IHave_MaxLengthOf = Not pExpressions.MaxLengthOf(CDbl(pNotHaveVal), Value, True)
End Function

Private Function IHave_MinLengthOf(Value As Double) As Boolean
    IHave_MinLengthOf = Not pExpressions.MinLengthOf(CDbl(pNotHaveVal), Value, True)
End Function

Private Function IHave_LengthBetween(lowVal As Double, highVal As Double) As Boolean
    IHave_LengthBetween = Not pExpressions.LengthBetween(CDbl(pNotHaveVal), lowVal, highVal, True)
End Function
