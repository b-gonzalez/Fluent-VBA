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
