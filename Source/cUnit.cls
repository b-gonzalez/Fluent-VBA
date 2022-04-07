VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pShould As cShould
Private pShouldSet As ISetExpressionAndValue
Private pShouldNot As cShouldNot
Private pShouldNotSet As ISetExpressionAndValue
Private pExpressions As cExpressions
Private pMeta As cMeta
Private pUnitVal As Variant
Private pTestCount As Long

Private Sub Class_Initialize()
    Set pExpressions = New cExpressions
End Sub

Public Property Set TestValue(value As Variant)
    pTestCount = pTestCount + 1
    Set pExpressions.TestValue = value
    Set pUnitVal = pExpressions.TestValue
End Property

Public Property Let TestValue(value As Variant)
    pTestCount = pTestCount + 1
    pExpressions.TestValue = value
    pUnitVal = pExpressions.TestValue
End Property

Public Property Get TestValue() As Variant
    If IsObject(pUnitVal) Then
        Set TestValue = pUnitVal
    Else
        TestValue = pUnitVal
    End If
End Property

Public Property Get Should() As IShould
    Set pShould = getShould
    Set Should = pShould
End Property

Public Property Get ShouldNot() As IShould
    Set pShouldNot = getShouldNot
    Set ShouldNot = pShouldNot
End Property

Public Property Set Meta(value As cMeta)
    Set pMeta = value
    Set pExpressions.setMeta = pMeta
End Property

Public Property Get Meta() As cMeta
    Set Meta = pMeta
End Property

Private Function getShould() As cShould
    Dim tempShould As ISetExpressionAndValue
    
    If tempShould Is Nothing Then
        Set tempShould = New cShould
    End If
    
    Set tempShould.setExpr = pExpressions
    
    If IsObject(pUnitVal) Then
        Set tempShould.TestValue = pUnitVal
    Else
        tempShould.TestValue = pUnitVal
    End If
    Set getShould = tempShould
End Function

Private Function getShouldNot() As cShouldNot
    Dim tempShouldNot As ISetExpressionAndValue
    
    If tempShouldNot Is Nothing Then
        Set tempShouldNot = New cShouldNot
    End If
    
    Set tempShouldNot.setExpr = pExpressions
    
    If IsObject(pUnitVal) Then
        Set tempShouldNot.TestValue = pUnitVal
    Else
        tempShouldNot.TestValue = pUnitVal
    End If
    
    Set getShouldNot = tempShouldNot
End Function

Private Sub Class_Terminate()
    If pMeta.PrintTotalTests = True Then
        If pTestCount = 1 Then
            Debug.Print pTestCount & " test completed"
        ElseIf pTestCount > 1 Then
            Debug.Print pTestCount & " tests completed"
        End If
    End If
End Sub
