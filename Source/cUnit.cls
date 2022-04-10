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

Implements IMeta
Implements ISetExpressionAndValue

Private pShould As cShould
Private pShouldSet As ISetExpressionAndValue
Private pShouldNot As cShouldNot
Private pShouldNotSet As ISetExpressionAndValue
Private pExpressions As cExpressions
Private pMeta As cMeta
Private pUnitVal As Variant
Private pTestCount As Long
Private pPrintSettings As cPrintSettings

Private Sub Class_Initialize()
    'Set pExpressions = New cExpressions
End Sub

Private Property Set ISetExpressionAndValue_TestValue(value As Variant)
    pTestCount = pTestCount + 1
    Set pExpressions.TestValue = value
    Set pUnitVal = pExpressions.TestValue
End Property

Private Property Let ISetExpressionAndValue_TestValue(value As Variant)
    pTestCount = pTestCount + 1
    pExpressions.TestValue = value
    pUnitVal = pExpressions.TestValue
End Property

'Private Property Get ISetExpressionAndValue_TestValue() As Variant
'    If IsObject(pUnitVal) Then
'        Set ISetExpressionAndValue_TestValue = pUnitVal
'    Else
'        ISetExpressionAndValue_TestValue = pUnitVal
'    End If
'End Property


Private Property Set ISetExpressionAndValue_setExpr(value As cExpressions)
    Set pExpressions = value
End Property

Public Property Get Should() As IShould
    Set pShould = getShould
    Set Should = pShould
End Property

Public Property Get ShouldNot() As IShould
    Set pShouldNot = getShouldNot
    Set ShouldNot = pShouldNot
End Property

Private Property Set IMeta_Meta(value As cMeta)
    Set pMeta = value
    Set pExpressions.setMeta = pMeta
End Property

Private Property Get IMeta_Meta() As cMeta
    Set IMeta_Meta = pMeta
End Property

Private Property Set IMeta_setPrintSettings(value As cPrintSettings)
    Set pPrintSettings = value
    Set pExpressions.setPrintSettings = pPrintSettings
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
    If pMeta.PrintSettings.PrintTotalTests = True Then
        If pTestCount = 1 Then
            Debug.Print pTestCount & " test completed"
        ElseIf pTestCount > 1 Then
            Debug.Print pTestCount & " tests completed"
        End If
    End If
End Sub
