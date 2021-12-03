VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExpressions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pTestValue As Variant
Private pMeta As cMeta

Public Property Let TestValue(value As Variant)
    pTestValue = value
End Property

Public Property Get TestValue() As Variant
    TestValue = pTestValue
End Property

Public Property Set setMeta(value As cMeta)
    Set pMeta = value
End Property

Public Function GreaterThan(OrigVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    GreaterThan = (OrigVal > NewVal)
    Call checkIfPrint(GreaterThan, negateValue)
End Function

Public Function GreaterThanOrEqualTo(OrigVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    GreaterThanOrEqualTo = (OrigVal >= NewVal)
    Call checkIfPrint(GreaterThanOrEqualTo, negateValue)
End Function

Public Function LessThan(OrigVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    LessThan = (OrigVal < NewVal)
    Call checkIfPrint(LessThan, negateValue)
End Function

Public Function LessThanOrEqualTo(OrigVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    LessThanOrEqualTo = (OrigVal <= NewVal)
    Call checkIfPrint(LessThanOrEqualTo, negateValue)
End Function

Public Function EqualTo(OrigVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    EqualTo = (OrigVal = NewVal)
    Call checkIfPrint(EqualTo, negateValue)
End Function

Public Function Contain(OrigVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    If OrigVal Like "*" & NewVal & "*" Then
        Contain = True
    End If
    Call checkIfPrint(Contain, negateValue)
End Function

Public Function StartWith(OrigVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    valLength = Len(NewVal)
    If Left(OrigVal, valLength) = CStr(NewVal) Then
        StartWith = True
    End If
    Call checkIfPrint(StartWith, negateValue)
End Function

Public Function EndWith(OrigVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    valLength = Len(NewVal)
    If Right(OrigVal, valLength) = CStr(NewVal) Then
        EndWith = True
    End If
    Call checkIfPrint(EndWith, negateValue)
End Function

Public Function LengthOf(OrigVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    LengthOf = (Len(CStr(OrigVal)) = NewVal)
    Call checkIfPrint(LengthOf, negateValue)
End Function

Public Function MaxLengthOf(OrigVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    MaxLengthOf = (Len(CStr(OrigVal)) <= NewVal)
    Call checkIfPrint(MaxLengthOf, negateValue)
End Function

Public Function MinLengthOf(OrigVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    MinLengthOf = (Len(CStr(OrigVal)) >= NewVal)
    Call checkIfPrint(MinLengthOf, negateValue)
End Function

Private Sub checkIfPrint(ByVal value As Boolean, Optional negateValue As Boolean = False)
    If pMeta.PrintResults Then
        If negateValue Then
            value = Not value
        End If
        printEval (value)
    End If
End Sub

Private Sub printEval(ByVal value As Boolean)
    Dim Result As String
    Dim TestPassed As Boolean
    
    Result = ""
    TestPassed = value

    If TestPassed Then
        Result = "Passed"
        If pMeta.TestName <> Empty Then
            Debug.Print pMeta.TestName & Result
        Else
            Debug.Print "Passed: " & Result
        End If
    Else
        Result = "Failed"
        If pMeta.TestName <> Empty Then
            Debug.Print pMeta.TestName & Result
        Else
            Debug.Print "Failed: " & Result
        End If
    End If
End Sub

