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

Private pPrintResults As Boolean
Private pTestName As String
Private pPrintTotalTests As Boolean
Private pApproximateEqual As Boolean
Private pEpsilon As Double

Public Property Let Epsilon(value As Double)
    pEpsilon = value
End Property

Public Property Get Epsilon() As Double
    Epsilon = pEpsilon
End Property

Public Property Let ApproximateEqual(value As Boolean)
    pApproximateEqual = value
End Property

Public Property Get ApproximateEqual() As Boolean
    ApproximateEqual = pApproximateEqual
End Property

Public Property Let TestName(value As String)
    pTestName = value
End Property

Public Property Get TestName() As String
    TestName = pTestName
End Property

Public Property Let PrintResults(value As Boolean)
    pPrintResults = value
End Property

Public Property Get PrintResults() As Boolean
    PrintResults = pPrintResults
End Property

Public Property Let PrintTotalTests(value As Boolean)
    pPrintTotalTests = value
End Property

Public Property Get PrintTotalTests() As Boolean
    PrintTotalTests = pPrintTotalTests
End Property

Private Sub Class_Initialize()
    Me.Epsilon = 0.000001
End Sub
