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
Private pPrintSettings As IPrintSettings

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

Public Property Get PrintSettings() As cPrintSettings
    Set PrintSettings = pPrintSettings
End Property

Public Property Let TestName(value As String)
    pTestName = value
    pPrintSettings.TestName = pTestName
End Property

Public Property Get TestName() As String
    TestName = pPrintSettings.TestName
End Property

Private Sub Class_Initialize()
    Me.Epsilon = defaultEpsilon
    Set pPrintSettings = New cPrintSettings
End Sub
