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

Public Function GreaterThan(OrigVal As Variant, NewVal As Variant, Optional NegateValue As Boolean = False) As Boolean
    GreaterThan = (OrigVal > NewVal)
    If pMeta.PrintResults Then
        If NegateValue Then
            NegateValue = Not GreaterThan
            PrintEval (NegateValue)
        Else
            PrintEval (GreaterThan)
        End If
    End If
End Function

Public Function LessThan(OrigVal As Variant, NewVal As Variant, Optional NegateValue As Boolean = False) As Boolean
    LessThan = (OrigVal < NewVal)
    If pMeta.PrintResults Then
        If NegateValue Then
            NegateValue = Not LessThan
            PrintEval (NegateValue)
        Else
            PrintEval (LessThan)
        End If
    End If
End Function

Public Function EqualTo(OrigVal As Variant, NewVal As Variant, Optional NegateValue As Boolean = False) As Boolean
    EqualTo = (OrigVal = NewVal)
    If pMeta.PrintResults Then
        If NegateValue Then
            NegateValue = Not EqualTo
            PrintEval (NegateValue)
        Else
            PrintEval (EqualTo)
        End If
    End If
End Function

Public Function Contain(OrigVal As Variant, NewVal As Variant, Optional NegateValue As Boolean = False) As Boolean
    If OrigVal Like "*" & NewVal & "*" Then
        Contain = True
    End If
    If pMeta.PrintResults Then
        If NegateValue Then
            NegateValue = Not Contain
            PrintEval (NegateValue)
        Else
            PrintEval (Contain)
        End If
    End If
End Function

Public Function StartWith(OrigVal As Variant, NewVal As Variant, Optional NegateValue As Boolean = False) As Boolean
    Dim valLength As Long
    valLength = Len(NewVal)
    If Left(OrigVal, valLength) = CStr(NewVal) Then
        StartWith = True
    End If
    If pMeta.PrintResults Then
        If NegateValue Then
            NegateValue = Not StartWith
            PrintEval (NegateValue)
        Else
            PrintEval (StartWith)
        End If
    End If
End Function

Public Function EndWith(OrigVal As Variant, NewVal As Variant, Optional NegateValue As Boolean = False) As Boolean
    Dim valLength As Long
    valLength = Len(NewVal)
    If Right(OrigVal, valLength) = CStr(NewVal) Then
        EndWith = True
    End If
    If pMeta.PrintResults Then
        If NegateValue Then
            NegateValue = Not EndWith
            PrintEval (NegateValue)
        Else
            PrintEval (EndWith)
        End If
    End If
End Function

Public Function LengthOf(OrigVal As Double, NewVal As Double, Optional NegateValue As Boolean = False) As Boolean
    LengthOf = (Len(CStr(OrigVal)) = NewVal)
    If pMeta.PrintResults Then
        If NegateValue Then
            NegateValue = Not LengthOf
            PrintEval (NegateValue)
        Else
            PrintEval (LengthOf)
        End If
    End If
End Function

Public Function MaxLengthOf(OrigVal As Double, NewVal As Double, Optional NegateValue As Boolean = False) As Boolean
    MaxLengthOf = (Len(CStr(OrigVal)) <= NewVal)
    If pMeta.PrintResults Then
        If NegateValue Then
            NegateValue = Not MaxLengthOf
            PrintEval (NegateValue)
        Else
            PrintEval (MaxLengthOf)
        End If
    End If
End Function

Public Function MinLengthOf(OrigVal As Double, NewVal As Double, Optional NegateValue As Boolean = False) As Boolean
    MinLengthOf = (Len(CStr(OrigVal)) >= NewVal)
    If pMeta.PrintResults Then
        If NegateValue Then
            NegateValue = Not MinLengthOf
            PrintEval (NegateValue)
        Else
            PrintEval (MinLengthOf)
        End If
    End If
End Function

Friend Sub PrintEval(ByVal value As Boolean)
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
