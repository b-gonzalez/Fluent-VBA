VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNotBe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IBe
Implements ISetExpressionAndValue

Private pNotBeValue As Variant
Private pBe As IBe
Private pExpressions As cExpressions

Private Property Set ISetExpressionAndValue_setExpr(value As cExpressions)
    Set pExpressions = value
End Property

Private Property Set ISetExpressionAndValue_TestValue(value As Variant)
    Set pNotBeValue = value
End Property

Private Property Let ISetExpressionAndValue_TestValue(value As Variant)
    pNotBeValue = value
End Property

Private Function IBe_GreaterThan(value As Double) As Boolean
    IBe_GreaterThan = Not pExpressions.GreaterThan(CDbl(pNotBeValue), value, True)
End Function

Private Function IBe_LessThan(value As Double) As Boolean
    IBe_LessThan = Not pExpressions.LessThan(CDbl(pNotBeValue), value, True)
End Function

Private Function IBe_EqualTo(value As Variant) As Boolean
    IBe_EqualTo = Not pExpressions.EqualTo(pNotBeValue, value, True)
End Function

Private Function IBe_GreaterThanOrEqualTo(value As Double) As Boolean
    IBe_GreaterThanOrEqualTo = Not pExpressions.GreaterThanOrEqualTo(CDbl(pNotBeValue), value, True)
End Function

Private Function IBe_LessThanOrEqualTo(value As Double) As Boolean
    IBe_LessThanOrEqualTo = Not pExpressions.LessThanOrEqualTo(CDbl(pNotBeValue), value, True)
End Function

Private Function IBe_Something() As Boolean
    IBe_Something = Not pExpressions.Something(pNotBeValue, True)
End Function

Private Function IBe_Between(lowerVal As Double, higherVal As Double) As Boolean
    IBe_Between = Not pExpressions.Between(CDbl(pNotBeValue), lowerVal, higherVal, True)
End Function

Private Function IBe_OneOf(ParamArray vals() As Variant) As Boolean
    IBe_OneOf = Not pExpressions.OneOf(pNotBeValue, False, vals)
End Function

Private Function IBe_InDataStructure(dataStructure As Variant) As Boolean
    IBe_InDataStructure = Not pExpressions.InDataStructure(pNotBeValue, dataStructure, True)
End Function
