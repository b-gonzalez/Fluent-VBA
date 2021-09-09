VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNotHave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IHave
Implements ISetExpression

Private pNotHaveVal As Variant
Private pExpressions As cExpressions

Public Property Set ISetExpression_setExpr(value As cExpressions)
    Set pExpressions = value
    pNotHaveVal = pExpressions.TestValue
End Property

Public Function IHave_LengthOf(value As Double) As Boolean
    IHave_LengthOf = Not pExpressions.LengthOf(CDbl(pNotHaveVal), value, True)
End Function

Public Function IHave_MaxLengthOf(value As Double) As Boolean
    IHave_MaxLengthOf = Not pExpressions.MaxLengthOf(CDbl(pNotHaveVal), value, True)
End Function

Public Function IHave_MinLengthOf(value As Double) As Boolean
    IHave_MinLengthOf = Not pExpressions.MinLengthOf(CDbl(pNotHaveVal), value, True)
End Function


