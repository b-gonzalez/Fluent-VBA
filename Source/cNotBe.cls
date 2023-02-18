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

Private Property Set ISetExpressionAndValue_setExpr(ByVal Value As cExpressions)
    Set pExpressions = Value
End Property

Private Property Set ISetExpressionAndValue_testValue(ByVal Value As Variant)
    Set pNotBeValue = Value
End Property

Private Property Let ISetExpressionAndValue_testValue(ByVal Value As Variant)
    pNotBeValue = Value
End Property

Private Function IBe_GreaterThan(ByVal Value As Double) As Boolean
    IBe_GreaterThan = Not pExpressions.GreaterThan(CDbl(pNotBeValue), Value, True)
End Function

Private Function IBe_LessThan(ByVal Value As Double) As Boolean
    IBe_LessThan = Not pExpressions.LessThan(CDbl(pNotBeValue), Value, True)
End Function

Private Function IBe_EqualTo(ByVal Value As Variant) As Boolean
    IBe_EqualTo = Not pExpressions.EqualTo(pNotBeValue, Value, True)
End Function

Private Function IBe_GreaterThanOrEqualTo(ByVal Value As Double) As Boolean
    IBe_GreaterThanOrEqualTo = Not pExpressions.GreaterThanOrEqualTo(CDbl(pNotBeValue), Value, True)
End Function

Private Function IBe_LessThanOrEqualTo(ByVal Value As Double) As Boolean
    IBe_LessThanOrEqualTo = Not pExpressions.LessThanOrEqualTo(CDbl(pNotBeValue), Value, True)
End Function

Private Function IBe_Something() As Boolean
    IBe_Something = Not pExpressions.Something(pNotBeValue, True)
End Function

Private Function IBe_Between(ByVal LowerVal As Double, ByVal HigherVal As Double) As Boolean
    IBe_Between = Not pExpressions.Between(CDbl(pNotBeValue), LowerVal, HigherVal, True)
End Function

Private Function IBe_OneOf(ParamArray Vals() As Variant)
    IBe_OneOf = Not pExpressions.OneOf(pNotBeValue, False, Vals)
End Function

Private Function IBe_InDataStructure(ByVal DataStructure As Variant, Optional ByVal Method As flMethod = flRecursive) As Boolean
    IBe_InDataStructure = Not pExpressions.InDataStructure(pNotBeValue, DataStructure, Method, True)
End Function
