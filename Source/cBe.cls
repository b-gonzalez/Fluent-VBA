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

Private Property Set ISetExpressionAndValue_setExpr(ByVal value As cExpressions)
    Set pExpressions = value
End Property

Private Property Let ISetExpressionAndValue_testValue(ByVal value As Variant)
    pBeValue = value
End Property

Private Property Set ISetExpressionAndValue_testValue(ByVal value As Variant)
    Set pBeValue = value
End Property

Private Function IBe_GreaterThan(ByVal value As Double) As Boolean
    IBe_GreaterThan = pExpressions.GreaterThan(CDbl(pBeValue), value)
End Function

Private Function IBe_LessThan(ByVal value As Double) As Boolean
    IBe_LessThan = pExpressions.LessThan(CDbl(pBeValue), value)
End Function

Private Function IBe_EqualTo(ByVal value As Variant) As Boolean
    IBe_EqualTo = pExpressions.EqualTo(pBeValue, value, updateFluentPath:=True)
End Function

Private Function IBe_GreaterThanOrEqualTo(ByVal value As Double) As Boolean
    IBe_GreaterThanOrEqualTo = pExpressions.GreaterThanOrEqualTo(CDbl(pBeValue), value, updateFluentPath:=True)
End Function

Private Function IBe_LessThanOrEqualTo(ByVal value As Double) As Boolean
    IBe_LessThanOrEqualTo = pExpressions.LessThanOrEqualTo(CDbl(pBeValue), value, updateFluentPath:=True)
End Function

Private Function IBe_Something() As Boolean
    IBe_Something = pExpressions.Something(pBeValue)
End Function

Private Function IBe_Between(ByVal LowerVal As Double, ByVal HigherVal As Double) As Boolean
    IBe_Between = pExpressions.Between(CDbl(pBeValue), LowerVal, HigherVal)
End Function

Private Function IBe_OneOf(ParamArray Vals() As Variant)
    IBe_OneOf = pExpressions.OneOf(pBeValue, False, Vals)
End Function

Private Function IBe_InDataStructure(ByVal DataStructure As Variant, Optional ByVal method As flAlgorithm = flRecursive) As Boolean
    IBe_InDataStructure = pExpressions.InDataStructure(pBeValue, DataStructure, method, updateFluentPath:=True)
End Function

Private Function IBe_InDataStructures(ByVal method As flAlgorithm, ParamArray DataStructures() As Variant) As Boolean
    IBe_InDataStructures = pExpressions.InDataStructures(pBeValue, method, False, DataStructures)
End Function

Private Function IBe_Alphabetic() As Boolean
    IBe_Alphabetic = pExpressions.Alphabetic(pBeValue)
End Function

Private Function IBe_Alphanumeric() As Boolean
    IBe_Alphanumeric = pExpressions.Alphanumeric(pBeValue)
End Function

Private Function IBe_Numeric() As Boolean
    IBe_Numeric = pExpressions.Numeric(pBeValue)
End Function

Private Function IBe_Erroneous() As Boolean
    IBe_Erroneous = pExpressions.Erroneous(pBeValue)
End Function

Private Function IBe_IdenticalTo(ByVal value As Variant) As Boolean
'    Dim temp As Variant
'    temp = args
    IBe_IdenticalTo = pExpressions.IdenticalTo(pBeValue, False, value, True)
End Function
