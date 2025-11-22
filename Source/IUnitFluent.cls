VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUnitFluent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Meta() As cMeta
End Property

Public Property Get Printing() As cPrinting
End Property

Public Property Get TestingFunctions() As cTestingFunctions
End Property

Public Property Let testValue(ByVal value As Variant)
End Property

Public Property Set testValue(ByVal value As Variant)
End Property

Public Property Get testValue() As Variant
End Property

Public Property Get Should() As IShould
End Property

Public Property Get ShouldNot() As IShould
End Property
