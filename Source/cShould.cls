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
Implements ISetTestFuncsAndValue

Private pShouldVal As Variant
Private pBe As cBe
Private pBeSet As ISetTestFuncsAndValue
Private pHave As cHave
Private pHaveSet As ISetTestFuncsAndValue
Private pTestingFunctions As ITestingFunctions

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pShouldVal = pTestingFunctions.TestValue
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pShouldVal = pTestingFunctions.TestValue
End Property

Private Property Get IShould_have() As IHave
    If pHave Is Nothing Then
        Set pHave = New cHave
    End If
    
    Set pHaveSet = pHave
    Set pHaveSet.SetTestFunc = pTestingFunctions
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
    Set pBeSet.SetTestFunc = pTestingFunctions
    
    If IsObject(pShouldVal) Then
        Set pBeSet.TestValue = pShouldVal
    Else
        pBeSet.TestValue = pShouldVal
    End If
    
    Set IShould_be = pBeSet
End Property

Private Function IShould_contain(ByVal value As String) As Variant
    IShould_contain = pTestingFunctions.Contain(pShouldVal, value)
End Function

Private Function IShould_startWith(ByVal value As String) As Variant
    IShould_startWith = pTestingFunctions.StartWith(pShouldVal, value)
End Function

Private Function IShould_endWith(ByVal value As String) As Variant
    IShould_endWith = pTestingFunctions.EndWith(pShouldVal, value)
End Function

Private Function IShould_evaluateTo(ByVal value As Variant) As Variant
    IShould_evaluateTo = pTestingFunctions.EvaluateTo(pShouldVal, value)
End Function
