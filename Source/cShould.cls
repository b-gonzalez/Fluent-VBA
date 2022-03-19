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

Private Property Set ISetExpressionAndValue_setExpr(value As cExpressions)
    Set pExpressions = value
End Property

Private Property Let ISetExpressionAndValue_TestValue(value As Variant)
    pShouldVal = pExpressions.TestValue
End Property

Private Property Set ISetExpressionAndValue_TestValue(value As Variant)
    Set pShouldVal = pExpressions.TestValue
End Property

Private Property Get IShould_Have() As IHave
    If pHave Is Nothing Then
        Set pHave = New cHave
    End If
    Set pHaveSet = pHave
    Set pHaveSet.setExpr = pExpressions
    pHaveSet.TestValue = pShouldVal
    Set IShould_Have = pHaveSet
End Property

Private Property Get IShould_Be() As IBe
    If pBe Is Nothing Then
        Set pBe = New cBe
    End If
    Set pBeSet = pBe
    Set pBeSet.setExpr = pExpressions
    
    If IsObject(pShouldVal) Then
        Set pBeSet.TestValue = pShouldVal
    Else
        pBeSet.TestValue = pShouldVal
    End If
    
    Set IShould_Be = pBeSet
End Property

Private Function IShould_Contain(value As Variant) As Boolean
    IShould_Contain = pExpressions.Contain(pShouldVal, value)
End Function

Private Function IShould_StartWith(value As Variant) As Boolean
    IShould_StartWith = pExpressions.StartWith(pShouldVal, value)
End Function

Private Function IShould_EndWith(value As Variant) As Boolean
    IShould_EndWith = pExpressions.EndWith(pShouldVal, value)
End Function

