VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFluentOf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pUnit As cUnit
Private pUnitMeta As IMeta
Private pUnitTestVal As ISetExpressionAndValue
Private pUnitExpr As ISetExpressionAndValue
Private pMeta As cMeta
Private pFluentOfVal As Variant
Private pExpressions As cExpressions

Private Sub Class_Initialize()
    Set pExpressions = New cExpressions
    Set pMeta = New cMeta
    Set pUnit = New cUnit
        
    Set pUnitExpr = pUnit
    Set pUnitTestVal = pUnit
    Set pUnitMeta = pUnit
    Set pUnitExpr.setExpr = pExpressions
    Set pUnitMeta.Meta = pMeta
    Set pUnitMeta.setPrintSettings = pMeta.PrintSettings

End Sub

Public Property Get Meta() As cMeta
    Set Meta = pMeta
End Property

Public Function Of(value As Variant) As cUnit
    If IsObject(value) Then
        Set pFluentOfVal = value
        Set pUnitTestVal.TestValue = value
    Else
        pFluentOfVal = value
        pUnitTestVal.TestValue = value
    End If
    Set Of = pUnit
End Function
