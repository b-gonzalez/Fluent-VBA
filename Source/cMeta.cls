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

Private pTestName As String
Private pApproximateEqual As Boolean
Private pEpsilon As Double
Private pTestCount As Long
Private pCategory As String
Private pPrinting As IPrinting
Private pFluentPath As String
Private pTestResult As cTestResult

Private Const defaultEpsilon As Double = 0.000001

Public Property Let ApproximateEqual(Value As Boolean)
    pApproximateEqual = Value
End Property

Public Property Get ApproximateEqual() As Boolean
    ApproximateEqual = pApproximateEqual
End Property

Public Property Let Epsilon(Value As Double)
    pEpsilon = Value
End Property

Public Property Get Epsilon() As Double
    Epsilon = pEpsilon
End Property

Public Property Let TestCount(Value As Long)
    pTestCount = Value
    pPrinting.TestCount = pTestCount
End Property

Public Property Get TestCount() As Long
    TestCount = pPrinting.TestCount
End Property

Public Property Get Printing() As cPrinting
    Set Printing = pPrinting
End Property

Public Property Get TestResult() As cTestResult
    Set TestResult = pTestResult
End Property

Private Sub Class_Initialize()
    Me.Epsilon = defaultEpsilon
    Set pPrinting = New cPrinting
    Set pTestResult = New cTestResult
End Sub
