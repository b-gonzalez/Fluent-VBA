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
Implements ISetExpression

Private pNotBeValue As Variant
Private pBe As IBe
Private pExpressions As cExpressions

Public Property Set ISetExpression_setExpr(value As cExpressions)
    Set pExpressions = value
    pNotBeValue = pExpressions.TestValue
End Property

Public Function IBe_GreaterThan(value As Variant) As Boolean
    IBe_GreaterThan = Not pExpressions.GreaterThan(pNotBeValue, value, True)
End Function

Public Function IBe_LessThan(value As Variant) As Boolean
    IBe_LessThan = Not pExpressions.LessThan(pNotBeValue, value, True)
End Function

Public Function IBe_EqualTo(value As Variant) As Boolean
    IBe_EqualTo = Not pExpressions.EqualTo(pNotBeValue, value, True)
End Function
