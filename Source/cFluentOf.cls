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

Implements IFluentOf

Private pUnit As IUnit
Private pUnitShould As cUnit
Private pMeta As IMeta
Private pTestingFunctions As cTestingFunctions
Private pFluentOfVal As Variant

Private Property Get IFluentOf_Meta() As IMeta
    Set IFluentOf_Meta = pMeta
End Property

Private Function IFluentOf_Of(Optional ByVal All As Variant) As cUnit
    Dim tf As ITestingFunctions
    If Not Information.IsMissing(All) Or Information.IsArray(All) Then
        If Information.IsObject(All) Then
            Set pFluentOfVal = All
            Set pUnit.TestValue = pFluentOfVal
        Else
            pFluentOfVal = All
            pUnit.TestValue = pFluentOfVal
        End If
    Else
        Set tf = pUnit.TestingFunctions
        tf.TestValueSet = False
    End If
    
    Set IFluentOf_Of = pUnit
End Function

Private Sub Class_Initialize()
    Set pUnit = New cUnit
    Set pUnitShould = pUnit
    Set pTestingFunctions = pUnit.TestingFunctions
    Set pMeta = pUnit.Meta
End Sub
