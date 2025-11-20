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
Private pShouldFunctions As IShouldTestFuncs
Private pShouldNotVal As Variant
Private pUtilities As IUtilities

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
    Set pShouldFunctions = value
End Property

Private Property Let ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    pShouldNotVal = pTestingFunctions.testValue
End Property

Private Property Set ISetTestFuncsAndValue_testValue(ByVal value As Variant)
    Set pShouldNotVal = pTestingFunctions.testValue
End Property

Private Property Get IShould_have() As IHave
    If pNotHave Is Nothing Then
        Set pNotHave = New cNotHave
    End If
    
    Set pNotHaveSet = pNotHave
    Set pNotHaveSet.SetTestFunc = pTestingFunctions
    
    If VBA.Information.IsObject(pShouldNotVal) Then
        Set pNotHaveSet.testValue = pShouldNotVal
    Else
        pNotHaveSet.testValue = pShouldNotVal
    End If
    
    Set IShould_have = pNotHaveSet
End Property

Private Property Get IShould_be() As IBe
    If pNotBe Is Nothing Then
        Set pNotBe = New cNotBe
    End If
    Set pNotBeSet = pNotBe
    Set pNotBeSet.SetTestFunc = pTestingFunctions
    
    If VBA.Information.IsObject(pShouldNotVal) Then
        Set pNotBeSet.testValue = pShouldNotVal
    Else
        pNotBeSet.testValue = pShouldNotVal
    End If
    
    Set IShould_be = pNotBeSet
End Property

Private Function IShould_Contain(ByVal value As String) As Variant
    Dim result As Variant
    result = pShouldFunctions.Contain(pShouldNotVal, value, True)
    result = pUtilities.NegateIfBool(result)
    IShould_Contain = result
End Function

Private Function IShould_StartWith(ByVal value As String) As Variant
    Dim result As Variant
    result = pShouldFunctions.StartWith(pShouldNotVal, value, True)
    result = pUtilities.NegateIfBool(result)
    IShould_StartWith = result
End Function

Private Function IShould_EndWith(ByVal value As String) As Variant
    Dim result As Variant
    result = pShouldFunctions.EndWith(pShouldNotVal, value, True)
    result = pUtilities.NegateIfBool(result)
    IShould_EndWith = result
End Function

Private Function IShould_EvaluateTo(ByVal value As Variant) As Variant
    Dim result As Variant
    result = pShouldFunctions.EvaluateTo(pShouldNotVal, value, True)
    result = pUtilities.NegateIfBool(result)
    IShould_EvaluateTo = result
End Function

Private Sub Class_Initialize()
    Set pUtilities = New cUtilities
End Sub

Private Sub Class_Terminate()
    Set pNotBe = Nothing
    Set pNotBeSet = Nothing
    Set pNotHave = Nothing
    Set pNotHaveSet = Nothing
    Set pTestingFunctions = Nothing
    Set pShouldFunctions = Nothing
End Sub
