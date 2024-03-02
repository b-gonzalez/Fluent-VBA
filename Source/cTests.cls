VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event TestPassed(ByVal test As cTest)
Public Event TestFailed(ByVal test As cTest)
Public Event TestCompleted(ByVal test As cTest)
Public Event TestUnexpected(ByVal test As cTest)

Private pResult As Boolean
Private pCol As Collection
Private pCount As Long

Private Sub Class_Initialize()
    Set pCol = New Collection
    pCount = 0
End Sub

Public Property Get Result() As Boolean
    Result = pResult
End Property

Public Property Get Item(OneBasedIndex As Long) As cTest
Attribute Item.VB_UserMemId = 0
    Set Item = pCol(OneBasedIndex)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCol.[_NewEnum]
End Property

Public Property Get Count() As Long
    Count = pCount
End Property

Public Sub ResetCounter()
    pCount = 0
End Sub

Public Sub CheckTest(test As cTest)
    pResult = test.Result
    
    pCol.Add test
    
    pCount = pCount + 1
    
    RaiseEvent TestCompleted(test)
    
    If test.Result = True Then
        RaiseEvent TestPassed(test)
    ElseIf test.Result = False Then
        RaiseEvent TestFailed(test)
    Else
        RaiseEvent TestUnexpected(test)
    End If
End Sub
