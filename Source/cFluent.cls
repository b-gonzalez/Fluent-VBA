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

Private pUnit As cUnit
Private pUnitMeta As IUnit
Private pUnitExpr As ISetExpressionAndValue
Private pMeta As cMeta
Private pFluentVal As Variant
Private pExpressions As cExpressions

Private Sub Class_Initialize()
    Set pExpressions = New cExpressions
    Set pUnit = New cUnit
    Set pMeta = New cMeta
    
    Set pUnitExpr = pUnit
    Set pUnitExpr.setExpr = pExpressions
    Set pUnitMeta = pUnit
    Set pUnitMeta.Meta = pMeta
    Set pUnitMeta.setPrintSettings = pMeta.PrintSettings
    
End Sub

Public Property Get Should() As IShould
    Set Should = pUnit.Should
End Property

Public Property Get ShouldNot() As IShould
    Set ShouldNot = pUnit.ShouldNot
End Property

Public Property Get Meta() As cMeta
    Set Meta = pMeta
End Property

Public Property Set TestValue(value As Variant)
    Set pFluentVal = value
    Set pUnitExpr.TestValue = pFluentVal
End Property

Public Property Let TestValue(value As Variant)
    pFluentVal = value
    pUnitExpr.TestValue = pFluentVal
End Property

Public Property Get TestValue() As Variant
    If IsObject(pFluentVal) Then
        Set TestValue = pExpressions.TestValue
    Else
        TestValue = pExpressions.TestValue
    End If
End Property
