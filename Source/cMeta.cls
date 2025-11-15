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
Private pTests As ITests
Private pUtilities As IUtilities
Private pToString As IToString

Private Property Get IMeta_Printing() As IPrinting
    Set IMeta_Printing = pPrinting
End Property

Private Property Get IMeta_Tests() As ITests
    Set IMeta_Tests = pTests
End Property

Private Property Get IMeta_Utilities() As IUtilities
    Set IMeta_Utilities = pUtilities
End Property

Private Property Get IMeta_ToString() As IToString
    Set IMeta_ToString = pToString
End Property

Private Sub Class_Initialize()
    Set pPrinting = New cPrinting
    Set pTests = New cTests
    Set pUtilities = New cUtilities
    Set pToString = New cToString
End Sub

Private Sub Class_Terminate()
    Set pPrinting = Nothing
    Set pTests = Nothing
    Set pUtilities = Nothing
    Set pToString = Nothing
End Sub
