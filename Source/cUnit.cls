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

Private pShouldSet As ISetTestFuncsAndValue
Private pShouldNotSet As ISetTestFuncsAndValue
Private pTestingFunctions As ITestingFunctions
Private pMeta As IMeta
Private pPrinting As cPrinting
Private iPrint As IPrintExpr

Private pUnitVal As Variant
Private pTestCount As Long

Public Property Get Should() As IShould
    Set Should = getShould
End Property

Public Property Get ShouldNot() As IShould
    Set ShouldNot = getShouldNot
End Property

Private Property Get IUnit_meta() As cMeta
    Set IUnit_meta = pMeta
End Property

Private Property Get IUnit_printing() As cPrinting
    Set IUnit_printing = pPrinting
End Property

Private Property Get IUnit_TestingFunctions() As cTestingFunctions
    Set IUnit_TestingFunctions = pTestingFunctions
End Property

Private Property Let IUnit_testValue(ByVal value As Variant)
    pTestingFunctions.TestValue = value
    pUnitVal = pTestingFunctions.TestValue
    
    With pMeta.Tests.TestStrings
        If .CleanTestValueStr Or .CleanTestStrings Then
            pTestingFunctions.TestValueClean = value
        End If
    End With
End Property

Private Property Set IUnit_testValue(ByVal value As Variant)
    Set pTestingFunctions.TestValue = value
    Set pUnitVal = pTestingFunctions.TestValue
End Property

Private Property Get IUnit_testValue() As Variant
    If VBA.Information.IsObject(pUnitVal) Then
        Set IUnit_testValue = pUnitVal
    Else
        IUnit_testValue = pUnitVal
    End If
End Property

Private Function getShould() As cShould
    Dim tempShould As ISetTestFuncsAndValue
    
    If tempShould Is Nothing Then
        Set tempShould = New cShould
    End If
    
    Set tempShould.SetTestFunc = pTestingFunctions
    
    If VBA.Information.IsObject(pUnitVal) Then
        Set tempShould.TestValue = pUnitVal
    Else
        tempShould.TestValue = pUnitVal
    End If
    Set getShould = tempShould
End Function

Private Function getShouldNot() As cShouldNot
    Dim tempShouldNot As ISetTestFuncsAndValue
    
    If tempShouldNot Is Nothing Then
        Set tempShouldNot = New cShouldNot
    End If
    
    Set tempShouldNot.SetTestFunc = pTestingFunctions
    
    If VBA.Information.IsObject(pUnitVal) Then
        Set tempShouldNot.TestValue = pUnitVal
    Else
        tempShouldNot.TestValue = pUnitVal
    End If
    
    Set getShouldNot = tempShouldNot
End Function

Private Sub Class_Initialize()
    Set pTestingFunctions = New cTestingFunctions
    Set pMeta = New cMeta
    
    Set iPrint = pMeta.Printing
    
    Set iPrint.Expressions.setMeta = pMeta
    Set iPrint.Expressions.setTests = pMeta.Tests
    Set iPrint.Expressions.setPrinting = pMeta.Printing
    
    Set pTestingFunctions.Expressions = iPrint.Expressions
    Set pTestingFunctions.setMeta = pMeta
End Sub
