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

Private Property Set ISetExpressionAndValue_setExpr(ByVal value As cExpressions)
    Set pExpressions = value
End Property

Private Property Set ISetExpressionAndValue_testValue(ByVal value As Variant)
    Set pNotBeValue = value
End Property

Private Property Let ISetExpressionAndValue_testValue(ByVal value As Variant)
    pNotBeValue = value
End Property

Private Function IBe_GreaterThan(ByVal value As Double) As Boolean
    IBe_GreaterThan = Not pExpressions.GreaterThan(CDbl(pNotBeValue), value, True)
End Function

Private Function IBe_LessThan(ByVal value As Double) As Boolean
    IBe_LessThan = Not pExpressions.LessThan(CDbl(pNotBeValue), value, True)
End Function

Private Function IBe_EqualTo(ByVal value As Variant) As Boolean
    IBe_EqualTo = Not pExpressions.EqualTo(pNotBeValue, value, True)
End Function

Private Function IBe_GreaterThanOrEqualTo(ByVal value As Double) As Boolean
    IBe_GreaterThanOrEqualTo = Not pExpressions.GreaterThanOrEqualTo(CDbl(pNotBeValue), value, True)
End Function

Private Function IBe_LessThanOrEqualTo(ByVal value As Double) As Boolean
    IBe_LessThanOrEqualTo = Not pExpressions.LessThanOrEqualTo(CDbl(pNotBeValue), value, True)
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

Private Function IBe_Alphabetic() As Boolean
    IBe_Alphabetic = Not pExpressions.Alphabetic(pNotBeValue, True)
End Function

Private Function IBe_Alphanumeric() As Boolean
    IBe_Alphanumeric = Not pExpressions.Alphanumeric(pNotBeValue, True)
End Function

Private Function IBe_Numeric() As Boolean
    IBe_Numeric = Not pExpressions.Numeric(pNotBeValue, True)
End Function

Private Function IBe_Erroneous() As Boolean
    IBe_Erroneous = Not pExpressions.Erroneous(pNotBeValue, True)
End Function
