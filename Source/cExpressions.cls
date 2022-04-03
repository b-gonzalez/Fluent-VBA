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

Enum flMethod
    flRecursive
    flIterative
End Enum

Public Property Let TestValue(value As Variant)
    pTestValue = value
End Property

Public Property Set TestValue(value As Variant)
    Set pTestValue = value
End Property

Public Property Get TestValue() As Variant
    If IsObject(pTestValue) Then
        Set TestValue = pTestValue
    Else
        TestValue = pTestValue
    End If
End Property

Public Property Set setMeta(value As cMeta)
    Set pMeta = value
End Property

Public Function GreaterThan(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    GreaterThan = (origVal > NewVal)
    Call checkIfPrint(GreaterThan, negateValue)
End Function

Public Function GreaterThanOrEqualTo(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    GreaterThanOrEqualTo = (origVal >= NewVal)
    Call checkIfPrint(GreaterThanOrEqualTo, negateValue)
End Function

Public Function LessThan(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    LessThan = (origVal < NewVal)
    Call checkIfPrint(LessThan, negateValue)
End Function

Public Function LessThanOrEqualTo(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    LessThanOrEqualTo = (origVal <= NewVal)
    Call checkIfPrint(LessThanOrEqualTo, negateValue)
End Function

Public Function EqualTo(origVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean

    Dim tempBool As Boolean
    
    tempBool = (origVal = NewVal)
    
    If pMeta.ApproximateEqual Then
        If TypeName(origVal) = TypeName(NewVal) Then
            tempBool = (origVal = NewVal)
        ElseIf TypeName(origVal) <> TypeName(NewVal) Then
            If IsNumeric(origVal) And IsNumeric(NewVal) Then
                tempBool = (Abs(origVal - NewVal) <= pMeta.Epsilon)
            ElseIf VarType(origVal) = vbString Or VarType(NewVal) = vbString Then
                tempBool = (StrComp(CStr(origVal), CStr(NewVal), vbTextCompare) = 0)
            End If
        End If
    End If
    
    EqualTo = tempBool

    Call checkIfPrint(EqualTo, negateValue)
End Function

Public Function Contain(origVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    If origVal Like "*" & NewVal & "*" Then
        Contain = True
    End If
    Call checkIfPrint(Contain, negateValue)
End Function

Public Function StartWith(origVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    valLength = Len(NewVal)
    If Left(origVal, valLength) = CStr(NewVal) Then
        StartWith = True
    End If
    Call checkIfPrint(StartWith, negateValue)
End Function

Public Function EndWith(origVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    valLength = Len(NewVal)
    If Right(origVal, valLength) = CStr(NewVal) Then
        EndWith = True
    End If
    Call checkIfPrint(EndWith, negateValue)
End Function

Public Function LengthOf(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    LengthOf = (Len(CStr(origVal)) = NewVal)
    Call checkIfPrint(LengthOf, negateValue)
End Function

Public Function MaxLengthOf(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    MaxLengthOf = (Len(CStr(origVal)) <= NewVal)
    Call checkIfPrint(MaxLengthOf, negateValue)
End Function

Public Function MinLengthOf(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    MinLengthOf = (Len(CStr(origVal)) >= NewVal)
    Call checkIfPrint(MinLengthOf, negateValue)
End Function

Public Function Something(origVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Something = (Not origVal Is Nothing)
    Call checkIfPrint(Something, negateValue)
End Function

Public Function Between(origVal As Double, lowerVal As Double, higherVal As Double, Optional negateValue As Boolean = False) As Boolean
    Between = (GreaterThanOrEqualTo(origVal, lowerVal) And LessThanOrEqualTo(origVal, higherVal))
    Call checkIfPrint(Between, negateValue)
End Function

Public Function OneOf(origVal As Variant, negateValue As Boolean, ParamArray args() As Variant)
    Dim i As Long
    Dim j As Long
    Dim tempBool As Boolean
    Dim arr As Variant
    
    arr = args(0)
    tempBool = False
    
    For i = LBound(arr) To UBound(arr)
        If IsObject(arr(i)) Then
            If TypeName(arr(i)) = TypeName(origVal) Then
                tempBool = True
            End If
        ElseIf arr(i) = origVal Then
            tempBool = True
            Exit For
        End If
    Next i
    
    OneOf = tempBool
    Call checkIfPrint(OneOf, negateValue)
End Function

Public Function InDataStructure(origVal As Variant, dataStructure As Variant, Optional method As flMethod = flRecursive, Optional negateValue As Boolean = False) As Boolean
    Dim elem As Variant
    Dim col As Collection
    Dim tempBool As Boolean
    
    tempBool = False
    
    If isDataStructure(dataStructure) Then
        If method = flRecursive Then
            Set col = getNestedElementsRecur(dataStructure)
        ElseIf method = flIterative Then
            Set col = getNestedElementsIter(dataStructure)
        End If
    End If
    
    For Each elem In col
        If origVal = elem Then
            tempBool = True
            Exit For
        End If
    Next elem
    
    InDataStructure = tempBool
    
    Call checkIfPrint(InDataStructure, negateValue)
    
End Function

Private Function isDataStructure(ByVal dataStructure As Variant)
    Dim dsTypeName As String
    Dim tempBool As Boolean
    
    dsTypeName = TypeName(dataStructure)
    tempBool = False
    
    Select Case True
        Case IsArray(dataStructure)
            tempBool = True
        Case TypeName(dataStructure) = "Collection", TypeName(dataStructure) = "Dictionary", TypeName(dataStructure) = "ArrayList"
            tempBool = True
    End Select
    
    isDataStructure = tempBool
End Function

Private Function getNestedElementsRecur(dataStructure As Variant, Optional ByRef col As Collection = Nothing) As Collection
    Dim elem As Variant
    Dim tempDataStructure As Variant
    
    If col Is Nothing Then
        Set col = New Collection
    End If
    
    If isDataStructure(dataStructure) Then
        If TypeName(dataStructure) = "Dictionary" Then
            tempDataStructure = dataStructure.Items
        Else
            If IsObject(dataStructure) Then
                Set tempDataStructure = dataStructure
            Else
                tempDataStructure = dataStructure
            End If
        End If
        
        For Each elem In tempDataStructure
            If isDataStructure(elem) Then
                Set col = getNestedElementsRecur(elem, col)
            Else
                col.Add elem
            End If
        Next elem
    Else
        col.Add elem
    End If
    
    Set getNestedElementsRecur = col
End Function

Private Function getNestedElementsIter(dataStructure As Variant) As Collection
    Dim elem As Variant
    Dim col As Collection
    Dim tempCol As Collection
    Dim tempDataStructure As Variant
    Dim i As Long

    If col Is Nothing Then
        Set col = New Collection
        Set tempCol = New Collection
    End If

    If TypeName(dataStructure) = "Dictionary" Then
        tempDataStructure = dataStructure.Items
    Else
        If IsObject(dataStructure) Then
            Set tempDataStructure = dataStructure
        Else
            tempDataStructure = dataStructure
        End If
    End If

    For Each elem In tempDataStructure
        If isDataStructure(elem) Then
            tempCol.Add elem
        Else
            col.Add elem
        End If
    Next elem

    Do While tempCol.Count > 0
        For i = 1 To tempCol.Count
            If isDataStructure(tempCol(i)) Then
                For Each elem In tempCol(i)
                    tempCol.Add elem
                Next elem
                tempCol.Remove i
                Exit For
            Else
                col.Add tempCol(i)
                tempCol.Remove i
                Exit For
            End If
        Next i
    Loop

    Set getNestedElementsIter = col
End Function

Public Function LengthBetween(origVal As Double, lowVal As Double, highVal As Double, Optional negateValue As Boolean = False)
    Dim tempOrigVal As Double
    tempOrigVal = Len(CStr(origVal))
    LengthBetween = (MinLengthOf(tempOrigVal, lowVal) And MaxLengthOf(tempOrigVal, highVal))
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
