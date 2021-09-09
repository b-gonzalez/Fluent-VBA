VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShould"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IShould
Implements ISetExpression

Private pShouldVal As Variant
Private pBe As cBe
Private pBeSet As ISetExpression
Private pHave As cHave
Private pHaveSet As ISetExpression
Private pExpressions As cExpressions

Public Property Set ISetExpression_setExpr(value As cExpressions)
    Set pExpressions = value
    pShouldVal = pExpressions.TestValue
End Property

Public Property Get IShould_Have() As IHave
    If pHave Is Nothing Then
        Set pHave = New cHave
    End If
    Set pHaveSet = pHave
    Set pHaveSet.setExpr = pExpressions
    Set IShould_Have = pHaveSet
End Property

Public Property Get IShould_Be() As IBe
    If pBe Is Nothing Then
        Set pBe = New cBe
    End If
    Set pBeSet = pBe
    Set pBeSet.setExpr = pExpressions
    Set IShould_Be = pBeSet
End Property

Public Function IShould_Contain(value As Variant) As Boolean
    IShould_Contain = pExpressions.Contain(pShouldVal, value)
End Function

Public Function IShould_StartWith(value As Variant) As Boolean
    IShould_StartWith = pExpressions.StartWith(pShouldVal, value)
End Function

Public Function IShould_EndWith(value As Variant) As Boolean
    IShould_EndWith = pExpressions.EndWith(pShouldVal, value)
End Function

