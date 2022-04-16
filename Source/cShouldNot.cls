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

Private Property Set ISetExpressionAndValue_setExpr(value As cExpressions)
    Set pExpressions = value
End Property

Private Property Let ISetExpressionAndValue_TestValue(value As Variant)
    pShouldNotVal = pExpressions.TestValue
End Property

Private Property Set ISetExpressionAndValue_TestValue(value As Variant)
    Set pShouldNotVal = pExpressions.TestValue
End Property

Private Property Get IShould_Have() As IHave
    If pNotHave Is Nothing Then
        Set pNotHave = New cNotHave
    End If
    Set pNotHaveSet = pNotHave
    Set pNotHaveSet.setExpr = pExpressions
    pNotHaveSet.TestValue = pShouldNotVal
    Set IShould_Have = pNotHaveSet
End Property

Private Property Get IShould_Be() As IBe
    If pNotBe Is Nothing Then
        Set pNotBe = New cNotBe
    End If
    Set pNotBeSet = pNotBe
    Set pNotBeSet.setExpr = pExpressions
    
    If IsObject(pShouldNotVal) Then
        Set pNotBeSet.TestValue = pShouldNotVal
    Else
        pNotBeSet.TestValue = pShouldNotVal
    End If
    
    Set IShould_Be = pNotBeSet
End Property

Private Function IShould_Contain(value As Variant) As Boolean
    IShould_Contain = Not pExpressions.Contain(pShouldNotVal, value, True)
End Function

Private Function IShould_StartWith(value As Variant) As Boolean
    IShould_StartWith = Not pExpressions.StartWith(pShouldNotVal, value, True)
End Function

Private Function IShould_EndWith(value As Variant) As Boolean
    IShould_EndWith = Not pExpressions.EndWith(pShouldNotVal, value, True)
End Function

Public Function IShould_EvaluateTo(value As Variant) As Boolean
    IShould_EvaluateTo = Not pExpressions.EvaluateTo(pShouldNotVal, value)
End Function
