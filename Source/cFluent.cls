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
Private pUnitMeta As IMeta
Private pMeta As cMeta
Private pFluentVal As Variant

Private Sub Class_Initialize()
    Set pUnit = New cUnit
    Set pMeta = New cMeta
    Set pUnitMeta = pUnit
    Set pUnitMeta.Meta = pMeta
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
    Set pUnit.TestValue = pFluentVal
End Property

Public Property Let TestValue(value As Variant)
    pFluentVal = value
    pUnit.TestValue = pFluentVal
End Property

Public Property Get TestValue() As Variant
    If IsObject(pFluentVal) Then
        Set TestValue = pFluentVal
    Else
        TestValue = pFluentVal
    End If
End Property
