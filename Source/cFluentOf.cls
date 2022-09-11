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

Private pUnit As IUnit
Private pUnitShould As cUnit
Private pMeta As cMeta
Private pExpressions As cExpressions
Private pFluentOfVal As Variant

Public Property Get Meta() As cMeta
    Set Meta = pMeta
End Property

Public Function Of(Value As Variant) As cUnit
    If IsObject(Value) Then
        Set pFluentOfVal = Value
        Set pUnit.TestValue = Value
    Else
        pFluentOfVal = Value
        pUnit.TestValue = Value
    End If
    Set Of = pUnit
End Function

Private Sub Class_Initialize()
    Set pUnit = New cUnit
    Set pUnitShould = pUnit
    Set pExpressions = pUnit.Expressiongs
    Set pMeta = pUnit.Meta
End Sub
