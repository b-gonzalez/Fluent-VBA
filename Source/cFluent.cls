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

Private pUnit As IUnitFluent
Private pMeta As IMeta
Private pTestingFunctions As cTestingFunctions
Private pFluentVal As Variant

Private Property Get IFluent_Should() As IShould
    Set IFluent_Should = pUnit.Should
End Property

Private Property Get IFluent_ShouldNot() As IShould
    Set IFluent_ShouldNot = pUnit.ShouldNot
End Property

Private Property Get IFluent_Meta() As IMeta
    Set IFluent_Meta = pMeta
End Property

Private Property Set IFluent_TestValue(ByVal All As Variant)
    Dim tf As ITestingFunctions
    
    If Not VBA.Information.IsMissing(All) Or VBA.Information.IsArray(All) Then
        Set pFluentVal = All
        Set pUnit.testValue = pFluentVal
    Else
        Set tf = pUnit.TestingFunctions
        tf.TestValueSet = False
    End If
End Property

Private Property Let IFluent_TestValue(ByVal All As Variant)
    Dim tf As ITestingFunctions
    
    If Not VBA.Information.IsMissing(All) Or VBA.Information.IsArray(All) Then
        pFluentVal = All
        pUnit.testValue = pFluentVal
    Else
        Set tf = pUnit.TestingFunctions
        tf.TestValueSet = False
    End If
End Property

Private Property Get IFluent_TestValue() As Variant
    If VBA.Information.IsObject(pFluentVal) Then
        Set IFluent_TestValue = pUnit.testValue
    Else
        IFluent_TestValue = pUnit.testValue
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

Public Property Set testValue(ByVal All As Variant)
    Set IFluent_TestValue = All
End Property

Public Property Let testValue(ByVal All As Variant)
    IFluent_TestValue = All
End Property

Public Property Get testValue() As Variant
    If VBA.Information.IsObject(pFluentVal) Then
        Set testValue = IFluent_TestValue
    Else
        testValue = IFluent_TestValue
    End If
End Property

Private Sub Class_Initialize()
    Set pUnit = New cUnit
    Set pTestingFunctions = pUnit.TestingFunctions
    Set pMeta = pUnit.Meta
End Sub

Private Sub Class_Terminate()
    Set pUnit = Nothing
    Set pTestingFunctions = Nothing
    Set pMeta = Nothing
End Sub
