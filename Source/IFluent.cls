VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFluent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Should() As IShould
End Property

Public Property Get ShouldNot() As IShould
End Property

Public Property Get Meta() As IMeta
End Property

Public Property Set TestValue(ByVal All As Variant)
End Property

Public Property Let TestValue(ByVal All As Variant)
End Property

Public Property Get TestValue() As Variant
End Property
