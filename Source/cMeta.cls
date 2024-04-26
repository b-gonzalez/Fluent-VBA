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

Private pTestName As String
Private pApproximateEqual As Boolean
Private pEpsilon As Double
Private pTestCount As Long
Private pCategory As String
Private pPrinting As IPrinting
Private pFluentPath As String
Private pTests As cTests

Private Const defaultEpsilon As Double = 0.000001

Private Property Let IMeta_ApproximateEqual(ByVal value As Boolean)
    pApproximateEqual = value
End Property

Private Property Get IMeta_ApproximateEqual() As Boolean
    IMeta_ApproximateEqual = pApproximateEqual
End Property

Private Property Let IMeta_Epsilon(ByVal value As Double)
    pEpsilon = value
End Property

Private Property Get IMeta_Epsilon() As Double
    IMeta_Epsilon = pEpsilon
End Property

Private Property Get IMeta_Printing() As IPrinting
    Set IMeta_Printing = pPrinting
End Property

Private Property Get IMeta_Tests() As cTests
    Set IMeta_Tests = pTests
End Property

Private Sub Class_Initialize()
    IMeta_Epsilon = defaultEpsilon
    Set pPrinting = New cPrinting
    Set pTests = New cTests
End Sub
