VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let ApproximateEqual(ByVal value As Boolean)
End Property

Public Property Get ApproximateEqual() As Boolean
End Property

Public Property Let Epsilon(ByVal value As Double)
End Property

Public Property Get Epsilon() As Double
End Property

Public Property Get Printing() As IPrinting
End Property

Public Property Get Tests() As cTests
End Property
