VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFluent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IFluent
Implements IFluentOf
Implements IFluentNoMeta

Private pShould As cShould
Private pShouldSet As ISetExpressionAndValue
Private pShouldNot As cShouldNot
Private pShouldNotSet As ISetExpressionAndValue
Private pExpressions As cExpressions
Private pMeta As cMeta
Private pMetaSet As ISetExpressionAndValue
Private pFluentVal As Variant

Private Sub Class_Initialize()
    Set pExpressions = New cExpressions
    Set pMeta = New cMeta
    Set pExpressions.setMeta = pMeta
End Sub

Private Property Let IFluent_TestValue(value As Variant)
    pExpressions.TestValue = value
    pFluentVal = pExpressions.TestValue
End Property

Private Property Get IFluent_TestValue() As Variant
    IFluent_TestValue = pFluentVal
End Property

Private Property Get IFluent_Should() As IShould
    Set pShould = getShould
    Set IFluent_Should = pShould
End Property

Private Property Get IFluent_ShouldNot() As IShould
    Set pShouldNot = getShouldNot
    Set IFluent_ShouldNot = pShouldNot
End Property

Private Property Get IFluent_Meta() As cMeta
    Set IFluent_Meta = pMeta
End Property

Private Property Get IFluentOf_Meta() As cMeta
    Set IFluentOf_Meta = pMeta
End Property

Private Function IFluentOf_Of(value As Variant) As IFluentNoMeta
    pExpressions.TestValue = value
    pFluentVal = pExpressions.TestValue
    Set IFluentOf_Of = Me
End Function

Private Property Get IFluentNoMeta_Should() As IShould
    Set pShould = getShould
    Set IFluentNoMeta_Should = pShould
End Property

Private Property Get IFluentNoMeta_ShouldNot() As IShould
    Set pShouldNot = getShouldNot
    Set IFluentNoMeta_ShouldNot = pShouldNot
End Property

Private Function getShould() As cShould
    Dim tempShould As ISetExpressionAndValue
    
    If tempShould Is Nothing Then
        Set tempShould = New cShould
    End If
    
    Set tempShould.setExpr = pExpressions
    tempShould.TestValue = pFluentVal
    Set getShould = tempShould
End Function

Private Function getShouldNot() As cShouldNot
    Dim tempShouldNot As ISetExpressionAndValue
    
    If tempShouldNot Is Nothing Then
        Set tempShouldNot = New cShouldNot
    End If
    
    Set tempShouldNot.setExpr = pExpressions
    tempShouldNot.TestValue = pFluentVal
    Set getShouldNot = tempShouldNot
End Function
