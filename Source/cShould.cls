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
Implements ISetExpressionAndValue

Private pShouldVal As Variant
Private pBe As cBe
Private pBeSet As ISetExpressionAndValue
Private pHave As cHave
Private pHaveSet As ISetExpressionAndValue
Private pExpressions As cExpressions

Private Property Set ISetExpressionAndValue_setExpr(ByVal value As cExpressions)
    Set pExpressions = value
End Property

Private Property Let ISetExpressionAndValue_testValue(ByVal value As Variant)
    pShouldVal = pExpressions.TestValue
End Property

Private Property Set ISetExpressionAndValue_testValue(ByVal value As Variant)
    Set pShouldVal = pExpressions.TestValue
End Property

Private Property Get IShould_have() As IHave
    If pHave Is Nothing Then
        Set pHave = New cHave
    End If
    
    Set pHaveSet = pHave
    Set pHaveSet.SetExpr = pExpressions
'    pHaveSet.TestValue = pShouldVal

    If IsObject(pShouldVal) Then
        Set pHaveSet.TestValue = pShouldVal
    Else
        pHaveSet.TestValue = pShouldVal
    End If
    
    Set IShould_have = pHaveSet
End Property

Private Property Get IShould_be() As IBe
    If pBe Is Nothing Then
        Set pBe = New cBe
    End If
    
    Set pBeSet = pBe
    Set pBeSet.SetExpr = pExpressions
    
    If IsObject(pShouldVal) Then
        Set pBeSet.TestValue = pShouldVal
    Else
        pBeSet.TestValue = pShouldVal
    End If
    
    Set IShould_be = pBeSet
End Property

Private Function IShould_contain(ByVal value As Variant) As Boolean
    IShould_contain = pExpressions.Contain(pShouldVal, value)
End Function

Private Function IShould_startWith(ByVal value As Variant) As Boolean
    IShould_startWith = pExpressions.StartWith(pShouldVal, value)
End Function

Private Function IShould_endWith(ByVal value As Variant) As Boolean
    IShould_endWith = pExpressions.EndWith(pShouldVal, value)
End Function

Private Function IShould_evaluateTo(ByVal value As Variant) As Boolean
    IShould_evaluateTo = pExpressions.EvaluateTo(pShouldVal, value)
End Function
