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

'Implements ITests

Public Event TestPassed(ByVal test As ITest)
Public Event TestFailed(ByVal test As ITest)
Public Event TestCompleted(ByVal test As ITest)
Public Event TestUnexpected(ByVal test As ITest)

Private pResult As Variant
Private pCol As Collection
Private pCount As Long
Private pTestDictCounter As Scripting.Dictionary
Private pAlgorithm As flAlgorithm
Private pToStrDev As Boolean

Private pApproximateEqual As Boolean
Private pEpsilon As Double
Private pDataStructures As Collection

Private Const defaultEpsilon As Double = 0.000001

Private Sub Class_Initialize()
    Set pCol = New Collection
    pCount = 0
    pAlgorithm = flRecursive
    
    pEpsilon = defaultEpsilon
    Set pDataStructures = New Collection
    
    pDataStructures.Add "Dictionary"
    pDataStructures.Add "Collection"
    pDataStructures.Add "ArrayList"
End Sub

Public Property Set TestDictCounter(value As Scripting.Dictionary)
    Set pTestDictCounter = value
End Property

Public Property Get TestDictCounter() As Scripting.Dictionary
    Set TestDictCounter = pTestDictCounter
End Property

Public Property Get Result() As Boolean
    Result = pResult
End Property

Public Property Get Item(OneBasedIndex As Long) As ITest
Attribute Item.VB_UserMemId = 0
    Set Item = pCol(OneBasedIndex)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = pCol.[_NewEnum]
End Property

Public Property Let Algorithm(value As flAlgorithm)
    pAlgorithm = value
End Property

Public Property Get Algorithm() As flAlgorithm
    Algorithm = pAlgorithm
End Property

Public Property Let ToStrDev(value As Boolean)
    pToStrDev = value
End Property

Public Property Get ToStrDev() As Boolean
    ToStrDev = pToStrDev
End Property

Public Property Get Count() As Long
    Count = pCount
End Property

Public Property Let ApproximateEqual(ByVal value As Boolean)
    pApproximateEqual = value
End Property

Public Property Get ApproximateEqual() As Boolean
    ApproximateEqual = pApproximateEqual
End Property

Public Property Let Epsilon(ByVal value As Double)
    pEpsilon = value
End Property

Public Property Get Epsilon() As Double
    Epsilon = pEpsilon
End Property

Public Property Get DataStructures() As Collection
    Set DataStructures = pDataStructures
End Property

Public Sub ResetCounter()
    pCount = 0
End Sub

Public Sub CheckTest(test As ITest)
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
