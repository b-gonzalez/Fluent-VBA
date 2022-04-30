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

Private Const defaultEpsilon As Double = 0.000001

Public Property Let ApproximateEqual(value As Boolean)
    pApproximateEqual = value
End Property

Public Property Get ApproximateEqual() As Boolean
    ApproximateEqual = pApproximateEqual
End Property

Public Property Let Epsilon(value As Double)
    pEpsilon = value
End Property

Public Property Get Epsilon() As Double
    Epsilon = pEpsilon
End Property

Public Property Get Printing() As cPrinting
    Set Printing = pPrinting
End Property

'Public Property Let testName(value As String)
'    pTestName = value
'    pPrintSettings.testName = pTestName
'End Property
'
'Public Property Get testName() As String
'    testName = pPrintSettings.testName
'End Property
'
'Public Property Let Category(value As String)
'    pCategory = value
'    pPrintSettings.Category = pCategory
'End Property
'
'Public Property Get Category() As String
'    Category = pPrintSettings.Category
'End Property

Public Property Let TestCount(value As Long)
    pTestCount = value
    pPrinting.TestCount = pTestCount
End Property

Public Property Get TestCount() As Long
    TestCount = pPrinting.TestCount
End Property

Private Sub Class_Initialize()
    Me.Epsilon = defaultEpsilon
    Set pPrinting = New cPrinting
End Sub
