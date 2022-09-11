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

Private Property Set ISetExpressionAndValue_setExpr(Value As cExpressions)
    Set pExpressions = Value
End Property

Private Property Let ISetExpressionAndValue_testValue(Value As Variant)
    pBeValue = Value
End Property

Private Property Set ISetExpressionAndValue_testValue(Value As Variant)
    Set pBeValue = Value
End Property

Private Function IBe_GreaterThan(Value As Double) As Boolean
    IBe_GreaterThan = pExpressions.GreaterThan(CDbl(pBeValue), Value)
End Function

Private Function IBe_LessThan(Value As Double) As Boolean
    IBe_LessThan = pExpressions.LessThan(CDbl(pBeValue), Value)
End Function

Private Function IBe_EqualTo(Value As Variant) As Boolean
    IBe_EqualTo = pExpressions.EqualTo(pBeValue, Value)
End Function

Private Function IBe_GreaterThanOrEqualTo(Value As Double) As Boolean
    IBe_GreaterThanOrEqualTo = pExpressions.GreaterThanOrEqualTo(CDbl(pBeValue), Value)
End Function

Private Function IBe_LessThanOrEqualTo(Value As Double) As Boolean
    IBe_LessThanOrEqualTo = pExpressions.LessThanOrEqualTo(CDbl(pBeValue), Value)
End Function

Private Function IBe_Something() As Boolean
    IBe_Something = pExpressions.Something(pBeValue)
End Function

Private Function IBe_Between(LowerVal As Double, HigherVal As Double) As Boolean
    IBe_Between = pExpressions.Between(CDbl(pBeValue), LowerVal, HigherVal)
End Function

Private Function IBe_OneOf(ParamArray Vals() As Variant)
    IBe_OneOf = pExpressions.OneOf(pBeValue, False, Vals)
End Function

Private Function IBe_InDataStructure(DataStructure As Variant, Optional Method As flMethod = flRecursive) As Boolean
    IBe_InDataStructure = pExpressions.InDataStructure(pBeValue, DataStructure, Method)
End Function
