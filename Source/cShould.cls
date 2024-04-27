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
Private pShouldFunctions As IShouldTestFuncs

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
    Set pShouldFunctions = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Dim tf As ITestingFunctions
    
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

Private Function IShould_Contain(ByVal value As String) As Variant
    IShould_Contain = pShouldFunctions.Contain(pShouldVal, value)
End Function

Private Function IShould_StartWith(ByVal value As String) As Variant
    IShould_StartWith = pShouldFunctions.StartWith(pShouldVal, value)
End Function

Private Function IShould_EndWith(ByVal value As String) As Variant
    IShould_EndWith = pShouldFunctions.EndWith(pShouldVal, value)
End Function

Private Function IShould_EvaluateTo(ByVal value As Variant) As Variant
    IShould_EvaluateTo = pShouldFunctions.EvaluateTo(pShouldVal, value)
End Function
