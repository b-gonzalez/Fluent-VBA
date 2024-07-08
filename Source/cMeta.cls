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
Private pUtilities As IUtilities

Private Property Get IMeta_Printing() As IPrinting
    Set IMeta_Printing = pPrinting
End Property

Private Property Get IMeta_Tests() As cTests
    Set IMeta_Tests = pTests
End Property

Public Property Get IMeta_Utilities() As IUtilities
    Set IMeta_Utilities = pUtilities
End Property

Private Sub Class_Initialize()
    Set pPrinting = New cPrinting
    Set pTests = New cTests
    Set pUtilities = New cUtilities
    
'    Set pUtilities.CleanStringDict = pTests.CleanStringDict
End Sub
