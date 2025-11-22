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

Implements IUnitFluent
Implements IUnitFluentOf

Private pShouldSet As ISetTestFuncsAndValue
Private pShouldNotSet As ISetTestFuncsAndValue
Private pTestingFunctions As ITestingFunctions
Private pMeta As IMeta
Private pPrinting As cPrinting
Private iPrint As IPrintExpr

Private pUnitVal As Variant
Private pTestCount As Long

Private Property Get IUnitFluent_meta() As cMeta
    Set IUnitFluent_meta = pMeta
End Property

Private Property Get IUnitFluent_printing() As cPrinting
    Set IUnitFluent_printing = pPrinting
End Property

Private Property Get IUnitFluent_TestingFunctions() As cTestingFunctions
    Set IUnitFluent_TestingFunctions = pTestingFunctions
End Property

Private Property Let IUnitFluent_testValue(ByVal value As Variant)
    pTestingFunctions.testValue = value
    pUnitVal = pTestingFunctions.testValue
    
    With pMeta.tests.TestStrings
        If .CleanTestValueStr Or .CleanTestStrings Then
            pTestingFunctions.TestValueClean = value
        End If
    End With
End Property

Private Property Set IUnitFluent_testValue(ByVal value As Variant)
    Set pTestingFunctions.testValue = value
    Set pUnitVal = pTestingFunctions.testValue
End Property

Private Property Get IUnitFluent_testValue() As Variant
    If VBA.Information.IsObject(pUnitVal) Then
        Set IUnitFluent_testValue = pUnitVal
    Else
        IUnitFluent_testValue = pUnitVal
    End If
End Property

Private Property Get IUnitFluent_Should() As IShould
    Set IUnitFluent_Should = getShould
End Property

Private Property Get IUnitFluent_ShouldNot() As IShould
    Set IUnitFluent_ShouldNot = getShouldNot
End Property

Private Property Get IUnitFluentOf_Should() As IShould
    Set IUnitFluentOf_Should = getShould
End Property

Private Property Get IUnitFluentOf_ShouldNot() As IShould
    Set IUnitFluentOf_ShouldNot = getShouldNot
End Property

Private Function getShould() As cShould
    Dim tempShould As ISetTestFuncsAndValue
    
    If tempShould Is Nothing Then
        Set tempShould = New cShould
    End If
    
    Set tempShould.SetTestFunc = pTestingFunctions
    
    If VBA.Information.IsObject(pUnitVal) Then
        Set tempShould.testValue = pUnitVal
    Else
        tempShould.testValue = pUnitVal
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
        Set tempShouldNot.testValue = pUnitVal
    Else
        tempShouldNot.testValue = pUnitVal
    End If
    
    Set getShouldNot = tempShouldNot
End Function

Private Sub Class_Initialize()
    Set pTestingFunctions = New cTestingFunctions
    Set pMeta = New cMeta
    
    Set iPrint = pMeta.Printing
    
    Set iPrint.Expressions.setMeta = pMeta
    Set iPrint.Expressions.setTests = pMeta.tests
    Set iPrint.Expressions.setPrinting = pMeta.Printing
    
    Set pTestingFunctions.Expressions = iPrint.Expressions
    Set pTestingFunctions.setMeta = pMeta
End Sub

Private Sub Class_Terminate()
    Set pMeta = Nothing
    
    Set iPrint.Expressions.setMeta = Nothing
    Set iPrint.Expressions.setTests = Nothing
    Set iPrint.Expressions.setPrinting = Nothing
    
    Set iPrint = Nothing
    
    Set pTestingFunctions.Expressions = Nothing
    Set pTestingFunctions.setMeta = Nothing
    
    Set pTestingFunctions = Nothing
End Sub
