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
Implements ISetExpressionAndValue

Private pNotBe As cNotBe
Private pNotBeSet As ISetExpressionAndValue
Private pNotHave As cNotHave
Private pNotHaveSet As ISetExpressionAndValue
Private pExpressions As cExpressions
Private pShouldNotVal As Variant

Private Property Set ISetExpressionAndValue_setExpr(ByVal value As cExpressions)
    Set pExpressions = value
End Property

Private Property Let ISetExpressionAndValue_testValue(ByVal value As Variant)
    pShouldNotVal = pExpressions.TestValue
End Property

Private Property Set ISetExpressionAndValue_testValue(ByVal value As Variant)
    Set pShouldNotVal = pExpressions.TestValue
End Property

Private Property Get IShould_have() As IHave
    If pNotHave Is Nothing Then
        Set pNotHave = New cNotHave
    End If
    
    Set pNotHaveSet = pNotHave
    Set pNotHaveSet.SetExpr = pExpressions
    
    If IsObject(pShouldNotVal) Then
        Set pNotHaveSet.TestValue = pShouldNotVal
    Else
        pNotHaveSet.TestValue = pShouldNotVal
    End If
    
    Set IShould_have = pNotHaveSet
End Property

Private Property Get IShould_be() As IBe
    If pNotBe Is Nothing Then
        Set pNotBe = New cNotBe
    End If
    Set pNotBeSet = pNotBe
    Set pNotBeSet.SetExpr = pExpressions
    
    If IsObject(pShouldNotVal) Then
        Set pNotBeSet.TestValue = pShouldNotVal
    Else
        pNotBeSet.TestValue = pShouldNotVal
    End If
    
    Set IShould_be = pNotBeSet
End Property

Private Function IShould_contain(ByVal value As Variant) As Variant
    IShould_contain = Not pExpressions.Contain(pShouldNotVal, value, True)
End Function

Private Function IShould_startWith(ByVal value As Variant) As Variant
    IShould_startWith = Not pExpressions.StartWith(pShouldNotVal, value, True)
End Function

Private Function IShould_endWith(ByVal value As Variant) As Variant
    IShould_endWith = Not pExpressions.EndWith(pShouldNotVal, value, True)
End Function

Public Function IShould_evaluateTo(ByVal value As Variant) As Variant
    IShould_evaluateTo = Not pExpressions.EvaluateTo(pShouldNotVal, value)
End Function
