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

Private pUnit As IUnit
Private pUnitShould As cUnit
Private pMeta As cMeta
Private pTestingFunctions As cTestingFunctions
Private pFluentVal As Variant

Private Property Get IFluent_Should() As IShould
    Set IFluent_Should = pUnitShould.Should
End Property

Private Property Get IFluent_ShouldNot() As IShould
    Set IFluent_ShouldNot = pUnitShould.ShouldNot
End Property

Private Property Get IFluent_Meta() As cMeta
    Set IFluent_Meta = pMeta
End Property

Private Property Set IFluent_TestValue(value As Variant)
    Set pFluentVal = value
    Set pUnit.TestValue = pFluentVal
End Property

Private Property Let IFluent_TestValue(ByVal value As Variant)
    pFluentVal = value
    pUnit.TestValue = pFluentVal
End Property

Private Property Get IFluent_TestValue() As Variant
    If IsObject(pFluentVal) Then
        Set IFluent_TestValue = pUnit.TestValue
    Else
        IFluent_TestValue = pUnit.TestValue
    End If
End Property

Private Sub Class_Initialize()
    Set pUnit = New cUnit
    Set pUnitShould = pUnit
    Set pTestingFunctions = pUnit.TestingFunctions
    Set pMeta = pUnit.Meta
End Sub
