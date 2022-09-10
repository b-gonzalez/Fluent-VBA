VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IBe
Implements ISetExpressionAndValue

Private pExpressions As cExpressions
Private pBeValue As Variant

Private Property Set ISetExpressionAndValue_setExpr(value As cExpressions)
    Set pExpressions = value
End Property

Private Property Let ISetExpressionAndValue_TestValue(value As Variant)
    pBeValue = value
End Property

Private Property Set ISetExpressionAndValue_TestValue(value As Variant)
    Set pBeValue = value
End Property

Private Function IBe_GreaterThan(value As Double) As Boolean
    IBe_GreaterThan = pExpressions.GreaterThan(CDbl(pBeValue), value)
End Function

Private Function IBe_LessThan(value As Double) As Boolean
    IBe_LessThan = pExpressions.LessThan(CDbl(pBeValue), value)
End Function

Private Function IBe_EqualTo(value As Variant) As Boolean
    IBe_EqualTo = pExpressions.EqualTo(pBeValue, value)
End Function

Private Function IBe_GreaterThanOrEqualTo(value As Double) As Boolean
    IBe_GreaterThanOrEqualTo = pExpressions.GreaterThanOrEqualTo(CDbl(pBeValue), value)
End Function

Private Function IBe_LessThanOrEqualTo(value As Double) As Boolean
    IBe_LessThanOrEqualTo = pExpressions.LessThanOrEqualTo(CDbl(pBeValue), value)
End Function

Private Function IBe_Something() As Boolean
    IBe_Something = pExpressions.Something(pBeValue)
End Function

Private Function IBe_Between(LowerVal As Double, HigherVal As Double) As Boolean
    IBe_Between = pExpressions.Between(CDbl(pBeValue), LowerVal, HigherVal)
End Function

Private Function IBe_OneOf(ParamArray vals() As Variant)
    IBe_OneOf = pExpressions.OneOf(pBeValue, False, vals)
End Function

Private Function IBe_InDataStructure(DataStructure As Variant, Optional Method As flMethod = flRecursive) As Boolean
    IBe_InDataStructure = pExpressions.InDataStructure(pBeValue, DataStructure, Method)
End Function
