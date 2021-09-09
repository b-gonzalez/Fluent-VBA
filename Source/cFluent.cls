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

Private pShould As cShould
Private pShouldSet As ISetExpression
Private pShouldNot As cShouldNot
Private pShouldNotSet As ISetExpression
Private pExpressions As cExpressions
Private pMeta As cMeta
Private pMetaSet As ISetExpression

Public Property Let TestValue(value As Variant)
    pExpressions.TestValue = value
End Property

Public Property Get TestValue() As Variant
    TestValue = pExpressions.TestValue
End Property

Public Property Get Should() As IShould
    If pShould Is Nothing Then
        Set pShould = New cShould
    End If
    Set pShouldSet = pShould
    Set pShouldSet.setExpr = pExpressions
    Set Should = pShouldSet
End Property

Public Property Get ShouldNot() As IShould
    If pShouldNot Is Nothing Then
        Set pShouldNot = New cShouldNot
    End If
    Set pShouldNotSet = pShouldNot
    Set pShouldNotSet.setExpr = pExpressions
    Set ShouldNot = pShouldNotSet
End Property

Public Property Get Meta() As cMeta
    Set Meta = pMeta
End Property

Private Sub Class_Initialize()
    Set pExpressions = New cExpressions
    Set pMeta = New cMeta
    Set pExpressions.setMeta = pMeta
End Sub
