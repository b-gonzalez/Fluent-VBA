VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFluentOf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pUnit As cUnit
Private pMeta As cMeta
Private pFluentOfVal As Variant

Private Sub Class_Initialize()
    Set pUnit = New cUnit
    Set pMeta = New cMeta
    Set pUnit.Meta = pMeta
End Sub

Public Property Get Meta() As cMeta
    Set Meta = pMeta
End Property

Public Function Of(value As Variant) As cUnit
    If IsObject(value) Then
        Set pFluentOfVal = value
        Set pUnit.TestValue = value
    Else
        pFluentOfVal = value
        pUnit.TestValue = value
    End If
    Set Of = pUnit
End Function
