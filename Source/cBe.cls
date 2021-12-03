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
Implements ISetExpression

Private pExpressions As cExpressions
Private pBeValue As Variant

Public Property Set ISetExpression_setExpr(value As cExpressions)
    Set pExpressions = value
    pBeValue = pExpressions.TestValue
End Property

Public Function IBe_GreaterThan(value As Double) As Boolean
    IBe_GreaterThan = pExpressions.GreaterThan(CDbl(pBeValue), value)
End Function

Public Function IBe_LessThan(value As Double) As Boolean
    IBe_LessThan = pExpressions.LessThan(CDbl(pBeValue), value)
End Function

Public Function IBe_EqualTo(value As Variant) As Boolean
    IBe_EqualTo = pExpressions.EqualTo(pBeValue, value)
End Function

Public Function IBe_GreaterThanOrEqualTo(value As Double) As Boolean
    IBe_GreaterThanOrEqualTo = pExpressions.GreaterThanOrEqualTo(CDbl(pBeValue), value)
End Function

Public Function IBe_LessThanOrEqualTo(value As Double) As Boolean
    IBe_LessThanOrEqualTo = pExpressions.LessThanOrEqualTo(CDbl(pBeValue), value)
End Function
