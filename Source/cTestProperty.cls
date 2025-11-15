VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTestProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const pTestCategory As Long = 1
Private Const pTestInfo As Long = 2
Private Const pTestName As Long = 3
Private Const pTestResultNumber As Long = 4

Public Property Get TestCategory() As Long
    TestCategory = pTestCategory
End Property

Public Property Get TestInfo() As Long
    TestInfo = pTestInfo
End Property

Public Property Get TestName() As Long
    TestName = pTestName
End Property

Public Property Get TestResultNumber() As Long
    TestResultNumber = pTestResultNumber
End Property
