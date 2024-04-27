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
Implements ISetTestFuncsAndValue

Private pNotBe As cNotBe
Private pNotBeSet As ISetTestFuncsAndValue
Private pNotHave As cNotHave
Private pNotHaveSet As ISetTestFuncsAndValue
Private pTestingFunctions As ITestingFunctions
Private pShouldNotVal As Variant

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pShouldNotVal = pTestingFunctions.TestValue
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pShouldNotVal = pTestingFunctions.TestValue
End Property

Private Property Get IShould_have() As IHave
    If pNotHave Is Nothing Then
        Set pNotHave = New cNotHave
    End If
    
    Set pNotHaveSet = pNotHave
    Set pNotHaveSet.SetTestFunc = pTestingFunctions
    
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
    Set pNotBeSet.SetTestFunc = pTestingFunctions
    
    If IsObject(pShouldNotVal) Then
        Set pNotBeSet.TestValue = pShouldNotVal
    Else
        pNotBeSet.TestValue = pShouldNotVal
    End If
    
    Set IShould_be = pNotBeSet
End Property

Private Function IShould_contain(ByVal value As String) As Variant
    IShould_contain = Not pTestingFunctions.Contain(pShouldNotVal, value, True)
End Function

Private Function IShould_startWith(ByVal value As String) As Variant
    IShould_startWith = Not pTestingFunctions.StartWith(pShouldNotVal, value, True)
End Function

Private Function IShould_endWith(ByVal value As String) As Variant
    IShould_endWith = Not pTestingFunctions.EndWith(pShouldNotVal, value, True)
End Function

Public Function IShould_evaluateTo(ByVal value As Variant) As Variant
    IShould_evaluateTo = Not pTestingFunctions.EvaluateTo(pShouldNotVal, value)
End Function
