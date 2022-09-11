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
Implements ISetExpressionAndValue

Private pExpressions As cExpressions
Private pHaveVal As Variant

Private Property Set ISetExpressionAndValue_setExpr(Value As cExpressions)
    Set pExpressions = Value
End Property

Private Property Let ISetExpressionAndValue_testValue(Value As Variant)
    pHaveVal = pExpressions.TestValue
End Property

Private Property Set ISetExpressionAndValue_testValue(Value As Variant)
    'intentionally left blank
End Property

Private Function IHave_LengthOf(Value As Double) As Boolean
    IHave_LengthOf = pExpressions.LengthOf(CDbl(pHaveVal), Value)
End Function

Private Function IHave_MaxLengthOf(Value As Double) As Boolean
    IHave_MaxLengthOf = pExpressions.MaxLengthOf(CDbl(pHaveVal), Value)
End Function

Private Function IHave_MinLengthOf(Value As Double) As Boolean
    IHave_MinLengthOf = pExpressions.MinLengthOf(CDbl(pHaveVal), Value)
End Function

Private Function IHave_LengthBetween(lowVal As Double, highVal As Double) As Boolean
    IHave_LengthBetween = pExpressions.LengthBetween(CDbl(pHaveVal), lowVal, highVal)
End Function
