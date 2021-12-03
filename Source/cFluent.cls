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
Private pShouldSet As ISetExpression
Private pShouldNot As cShouldNot
Private pShouldNotSet As ISetExpression
Private pExpressions As cExpressions
Private pMeta As cMeta
Private pMetaSet As ISetExpression

Private Sub Class_Initialize()
    Set pExpressions = New cExpressions
    Set pMeta = New cMeta
    Set pExpressions.setMeta = pMeta
End Sub

Public Property Let IFluent_TestValue(value As Variant)
    pExpressions.TestValue = value
End Property

Public Property Get IFluent_TestValue() As Variant
    IFluent_TestValue = pExpressions.TestValue
End Property

Public Property Get IFluent_Should() As IShould
    Set pShould = getShould
    Set IFluent_Should = pShould
End Property

Public Property Get IFluent_ShouldNot() As IShould
    Set pShouldNot = getShouldNot
    Set IFluent_ShouldNot = pShouldNot
End Property

Public Property Get IFluent_Meta() As cMeta
    Set IFluent_Meta = pMeta
End Property

Public Property Get IFluentOf_Meta() As cMeta
    Set IFluentOf_Meta = pMeta
End Property

Public Function IFluentOf_Of(value As Variant) As IFluentNoMeta
    pExpressions.TestValue = value
    Set IFluentOf_Of = Me
End Function

Public Property Get IFluentNoMeta_Should() As IShould
    Set pShould = getShould
    Set IFluentNoMeta_Should = pShould
End Property

Public Property Get IFluentNoMeta_ShouldNot() As IShould
    Set pShouldNot = getShouldNot
    Set IFluentNoMeta_ShouldNot = pShouldNot
End Property

Private Function getShould() As cShould
    Dim tempShould As ISetExpression
    
    If tempShould Is Nothing Then
        Set tempShould = New cShould
    End If
    
    Set tempShould.setExpr = pExpressions
    Set getShould = tempShould
End Function

Private Function getShouldNot() As cShouldNot
    Dim tempShouldNot As ISetExpression
    
    If tempShouldNot Is Nothing Then
        Set tempShouldNot = New cShouldNot
    End If
    
    Set tempShouldNot.setExpr = pExpressions
    Set getShouldNot = tempShouldNot
End Function
