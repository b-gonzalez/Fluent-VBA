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

Private Sub Class_Initialize()
    Set pUtilities = New cUtilities
End Sub

Private Property Set ISetTestFuncsAndValue_SetTestFunc(ByVal value As cTestingFunctions)
    Set pTestingFunctions = value
    Set pShouldFunctions = value
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
    
    If Information.IsObject(pShouldNotVal) Then
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
    
    If Information.IsObject(pShouldNotVal) Then
        Set pNotBeSet.TestValue = pShouldNotVal
    Else
        pNotBeSet.TestValue = pShouldNotVal
    End If
    
    Set IShould_be = pNotBeSet
End Property

Private Function IShould_Contain(ByVal value As String) As Variant
    Dim Result As Variant
    Result = pShouldFunctions.Contain(pShouldNotVal, value, True)
    Result = pUtilities.NegateIfBool(Result)
    IShould_Contain = Result
End Function

Private Function IShould_StartWith(ByVal value As String) As Variant
    Dim Result As Variant
    Result = pShouldFunctions.StartWith(pShouldNotVal, value, True)
    Result = pUtilities.NegateIfBool(Result)
    IShould_StartWith = Result
End Function

Private Function IShould_EndWith(ByVal value As String) As Variant
    Dim Result As Variant
    Result = pShouldFunctions.EndWith(pShouldNotVal, value, True)
    Result = pUtilities.NegateIfBool(Result)
    IShould_EndWith = Result
End Function

Private Function IShould_EvaluateTo(ByVal value As Variant) As Variant
    Dim Result As Variant
    Result = pShouldFunctions.EvaluateTo(pShouldNotVal, value, True)
    Result = pUtilities.NegateIfBool(Result)
    IShould_EvaluateTo = Result
End Function
