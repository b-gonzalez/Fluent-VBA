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

Private pUnit As IUnit
Private pUnitShould As cUnit
Private pMeta As cMeta
Private pExpressions As cExpressions
Private pFluentVal As Variant

Private Sub Class_Initialize()
    Set pUnit = New cUnit
    Set pUnitShould = pUnit
    Set pExpressions = pUnit.Expressiongs
    Set pMeta = pUnit.Meta
End Sub

Public Property Get Should() As IShould
    Set Should = pUnitShould.Should
End Property

Public Property Get ShouldNot() As IShould
    Set ShouldNot = pUnitShould.ShouldNot
End Property

Public Property Get Meta() As cMeta
    Set Meta = pMeta
End Property

Public Property Set TestValue(value As Variant)
    Set pFluentVal = value
    Set pUnit.TestValue = pFluentVal
End Property

Public Property Let TestValue(value As Variant)
    pFluentVal = value
    pUnit.TestValue = pFluentVal
End Property

Public Property Get TestValue() As Variant
    If IsObject(pFluentVal) Then
        Set TestValue = pUnit.TestValue
    Else
        TestValue = pUnit.TestValue
    End If
End Property
