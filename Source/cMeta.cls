VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IMeta

Private pPrinting As IPrinting
Private pTests As cTests

Private Sub Class_Initialize()
    Set pPrinting = New cPrinting
    Set pTests = New cTests
End Sub

Private Property Get IMeta_Printing() As IPrinting
    Set IMeta_Printing = pPrinting
End Property

Private Property Get IMeta_Tests() As cTests
    Set IMeta_Tests = pTests
End Property
