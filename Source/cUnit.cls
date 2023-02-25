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

Implements IUnit

Private pShould As cShould
Private pShouldSet As ISetExpressionAndValue
Private pShouldNot As cShouldNot
Private pShouldNotSet As ISetExpressionAndValue
Private pExpressions As cExpressions
Private pMeta As cMeta
Private pPrinting As cPrinting
Private iPrint As IPrinting

Private pUnitVal As Variant
Private pTestCount As Long

Public Property Get Should() As IShould
    Set pShould = getShould
    Set Should = pShould
End Property

Public Property Get ShouldNot() As IShould
    Set pShouldNot = getShouldNot
    Set ShouldNot = pShouldNot
End Property

Private Property Get IUnit_meta() As cMeta
    Set IUnit_meta = pMeta
End Property

Private Property Get IUnit_printing() As cPrinting
    Set IUnit_printing = pPrinting
End Property

Private Property Get IUnit_expressiongs() As cExpressions
    Set IUnit_expressiongs = pExpressions
End Property

Private Property Let IUnit_testValue(ByVal value As Variant)
    pExpressions.TestValue = value
    pUnitVal = pExpressions.TestValue
End Property

Private Property Set IUnit_testValue(ByVal value As Variant)
    Set pExpressions.TestValue = value
    Set pUnitVal = pExpressions.TestValue
End Property

Private Property Get IUnit_testValue() As Variant
    IUnit_testValue = pUnitVal
End Property

Private Function getShould() As cShould
    Dim tempShould As ISetExpressionAndValue
    
    If tempShould Is Nothing Then
        Set tempShould = New cShould
    End If
    
    Set tempShould.SetExpr = pExpressions
    
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
    
    Set tempShouldNot.SetExpr = pExpressions
    
    If IsObject(pUnitVal) Then
        Set tempShouldNot.TestValue = pUnitVal
    Else
        tempShouldNot.TestValue = pUnitVal
    End If
    
    Set getShouldNot = tempShouldNot
End Function

Private Sub Class_Initialize()
    Set pExpressions = New cExpressions
    Set pMeta = New cMeta
    
    Set pExpressions.setMeta = pMeta
    Set pExpressions.setPrinting = pMeta.Printing
    Set pExpressions.setTestResult = pMeta.TestResult
End Sub

