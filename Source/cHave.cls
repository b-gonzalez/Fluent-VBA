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

Private Property Set ISetExpressionAndValue_setExpr(value As cExpressions)
    Set pExpressions = value
End Property

Private Property Let ISetExpressionAndValue_TestValue(value As Variant)
    pHaveVal = pExpressions.TestValue
End Property

Private Function IHave_LengthOf(value As Double) As Boolean
    IHave_LengthOf = pExpressions.LengthOf(CDbl(pHaveVal), value)
End Function

Private Function IHave_MaxLengthOf(value As Double) As Boolean
    IHave_MaxLengthOf = pExpressions.MaxLengthOf(CDbl(pHaveVal), value)
End Function

Private Function IHave_MinLengthOf(value As Double) As Boolean
    IHave_MinLengthOf = pExpressions.MinLengthOf(CDbl(pHaveVal), value)
End Function
