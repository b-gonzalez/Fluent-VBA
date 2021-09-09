VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShouldNot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IShould
Implements ISetExpression

Private pNotBe As cNotBe
Private pNotBeSet As ISetExpression
Private pNotHave As cNotHave
Private pNotHaveSet As ISetExpression
Private pExpressions As cExpressions
Private pShouldNotVal As Variant


Public Property Set ISetExpression_setExpr(value As cExpressions)
    Set pExpressions = value
    pShouldNotVal = pExpressions.TestValue
End Property

Public Property Get IShould_Have() As IHave
    If pNotHave Is Nothing Then
        Set pNotHave = New cNotHave
    End If
    Set pNotHaveSet = pNotHave
    Set pNotHaveSet.setExpr = pExpressions
    Set IShould_Have = pNotHaveSet
End Property

Public Property Get IShould_Be() As IBe
    If pNotBe Is Nothing Then
        Set pNotBe = New cNotBe
    End If
    Set pNotBeSet = pNotBe
    Set pNotBeSet.setExpr = pExpressions
    Set IShould_Be = pNotBeSet
End Property

Public Function IShould_Contain(value As Variant) As Boolean
    IShould_Contain = Not pExpressions.Contain(pShouldNotVal, value, True)
End Function

Public Function IShould_StartWith(value As Variant) As Boolean
    IShould_StartWith = Not pExpressions.StartWith(pShouldNotVal, value, True)
End Function

Public Function IShould_EndWith(value As Variant) As Boolean
    IShould_EndWith = Not pExpressions.EndWith(pShouldNotVal, value, True)
End Function

