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
Private pMeta As IMeta
Private pTestingFunctions As cTestingFunctions
Private pFluentVal As Variant

Private Property Get IFluent_Should() As IShould
    Set IFluent_Should = pUnitShould.Should
End Property

Private Property Get IFluent_ShouldNot() As IShould
    Set IFluent_ShouldNot = pUnitShould.ShouldNot
End Property

Private Property Get IFluent_Meta() As IMeta
    Set IFluent_Meta = pMeta
End Property

Private Property Set IFluent_TestValue(ByVal All As Variant)
    Set pFluentVal = All
    Set pUnit.TestValue = pFluentVal
End Property

Private Property Let IFluent_TestValue(ByVal All As Variant)
    pFluentVal = All
    pUnit.TestValue = pFluentVal
End Property

Private Property Get IFluent_TestValue() As Variant
    If VBA.Information.IsObject(pFluentVal) Then
        Set IFluent_TestValue = pUnit.TestValue
    Else
        IFluent_TestValue = pUnit.TestValue
    End If
End Property

Public Property Get Should() As IShould
    Set Should = IFluent_Should
End Property

Public Property Get ShouldNot() As IShould
    Set ShouldNot = IFluent_ShouldNot
End Property

Public Property Get Meta() As IMeta
    Set Meta = IFluent_Meta
End Property

Public Property Set TestValue(ByVal All As Variant)
    Set IFluent_TestValue = All
End Property

Public Property Let TestValue(ByVal All As Variant)
    IFluent_TestValue = All
End Property

Public Property Get TestValue() As Variant
    If VBA.Information.IsObject(pFluentVal) Then
        Set TestValue = IFluent_TestValue
    Else
        TestValue = IFluent_TestValue
    End If
End Property



Private Sub Class_Initialize()
    Set pUnit = New cUnit
    Set pUnitShould = pUnit
    Set pTestingFunctions = pUnit.TestingFunctions
    Set pMeta = pUnit.Meta
End Sub
