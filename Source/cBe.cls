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

Public Function IBe_GreaterThan(value As Variant) As Boolean
    IBe_GreaterThan = pExpressions.GreaterThan(pBeValue, value)
End Function

Public Function IBe_LessThan(value As Variant) As Boolean
    IBe_LessThan = pExpressions.LessThan(pBeValue, value)
End Function

Public Function IBe_EqualTo(value As Variant) As Boolean
    IBe_EqualTo = pExpressions.EqualTo(pBeValue, value)
End Function

