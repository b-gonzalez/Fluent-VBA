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
Private pPrintSettings As IPrintSettings

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

Public Property Set setPrintSettings(value As IPrintSettings)
    Set pPrintSettings = value
End Property

Public Function InDataStructure(origVal As Variant, dataStructure As Variant, Optional method As flMethod = flRecursive, Optional negateValue As Boolean = False) As Boolean
    Dim elem As Variant
    Dim col As Collection
    Dim tempBool As Boolean
    Dim fluentPath As String
    
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
    
    fluentPath = CLng(InDataStructure) & " = " & pPrintSettings.TestName & " Should be in datastructure. "
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(InDataStructure, fluentPath, negateValue)
    
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

Public Function GreaterThan(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    
    GreaterThan = (origVal > NewVal)
    
    fluentPath = pPrintSettings.TestName & " Should be greater than. "
    fluentPath = CLng(GreaterThan) & " = " & fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(GreaterThan, fluentPath, negateValue)
    
End Function

Public Function GreaterThanOrEqualTo(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    
    GreaterThanOrEqualTo = (origVal >= NewVal)
    
    fluentPath = CLng(GreaterThanOrEqualTo) & " = " & pPrintSettings.TestName & " Should be greater than or equal to. "
    fluentPath = fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(GreaterThanOrEqualTo, fluentPath, negateValue)
End Function

Public Function LessThan(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    
    LessThan = (origVal < NewVal)
    
    fluentPath = CLng(LessThan) & " = " & pPrintSettings.TestName & " Should be less than. "
    fluentPath = fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(LessThan, fluentPath, negateValue)
End Function

Public Function LessThanOrEqualTo(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    
    LessThanOrEqualTo = (origVal <= NewVal)
    
    fluentPath = CLng(LessThanOrEqualTo) & " - " & pPrintSettings.TestName & " Should be less than or equal to. "
    fluentPath = fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(LessThanOrEqualTo, fluentPath, negateValue)
End Function

Public Function EqualTo(origVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim tempBool As Boolean
    Dim fluentPath As String
    
    
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
    
    fluentPath = CLng(EqualTo) & " = " & pPrintSettings.TestName & " Should be equal to. "
    fluentPath = fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(EqualTo, fluentPath, negateValue)
End Function

Public Function Contain(origVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    
    If origVal Like "*" & NewVal & "*" Then
        Contain = True
    End If
    
    fluentPath = CLng(Contain) & " = " & pPrintSettings.TestName & " Should contain. "
    fluentPath = fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(Contain, fluentPath, negateValue)
End Function

Public Function StartWith(origVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim fluentPath As String
    
    valLength = Len(NewVal)
    If Left(origVal, valLength) = CStr(NewVal) Then
        StartWith = True
    End If
    
    fluentPath = CLng(StartWith) & " = " & pPrintSettings.TestName & " Should start with. "
    fluentPath = fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(StartWith, fluentPath, negateValue)
End Function

Public Function EndWith(origVal As Variant, NewVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim valLength As Long
    Dim fluentPath As String
    
    valLength = Len(NewVal)
    If Right(origVal, valLength) = CStr(NewVal) Then
        EndWith = True
    End If
    
    fluentPath = CLng(EndWith) & " = " & pPrintSettings.TestName & " Should end with. "
    fluentPath = fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(EndWith, fluentPath, negateValue)
End Function

Public Function LengthOf(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    
    LengthOf = (Len(CStr(origVal)) = NewVal)
    
    fluentPath = CLng(LengthOf) & " = " & pPrintSettings.TestName & " Should have length of. "
    fluentPath = fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(LengthOf, fluentPath, negateValue)
End Function

Public Function MaxLengthOf(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    
    MaxLengthOf = (Len(CStr(origVal)) <= NewVal)
    
    fluentPath = CLng(MaxLengthOf) & " = " & pPrintSettings.TestName & " Should have max length of. "
    fluentPath = fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(MaxLengthOf, fluentPath, negateValue)
End Function

Public Function MinLengthOf(origVal As Double, NewVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    
    MinLengthOf = (Len(CStr(origVal)) >= NewVal)
    
    fluentPath = CLng(MinLengthOf) & " = " & pPrintSettings.TestName & " Should have min length of. "
    fluentPath = fluentPath & "Expected: " & origVal & "; Actual: " & NewVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(MinLengthOf, fluentPath, negateValue)
End Function

Public Function Something(origVal As Variant, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    
    Something = (Not origVal Is Nothing)
    
    fluentPath = CLng(Something) & " = " & pPrintSettings.TestName & " Should be equal something. "
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(Something, fluentPath, negateValue)
End Function

Public Function Between(origVal As Double, lowerVal As Double, higherVal As Double, Optional negateValue As Boolean = False) As Boolean
    Dim fluentPath As String
    
    Between = (GreaterThanOrEqualTo(origVal, lowerVal) And LessThanOrEqualTo(origVal, higherVal))
    
    fluentPath = CLng(Between) & " = " & pPrintSettings.TestName & " Should be between. "
    fluentPath = fluentPath & "Expected: Greater than " & lowerVal & " less than " & higherVal & "; Actual: " & origVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(Between, fluentPath, negateValue)
End Function

Public Function OneOf(origVal As Variant, negateValue As Boolean, ParamArray args() As Variant)
    Dim i As Long
    Dim j As Long
    Dim tempBool As Boolean
    Dim arr As Variant
    Dim fluentPath As String
    
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
    
    fluentPath = CLng(OneOf) & " = " & pPrintSettings.TestName & " Should be one of. "
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(OneOf, fluentPath, negateValue)
End Function

Public Function LengthBetween(origVal As Double, lowerVal As Double, higherVal As Double, Optional negateValue As Boolean = False)
    Dim tempOrigVal As Double
    Dim fluentPath As String
    
    tempOrigVal = Len(CStr(origVal))
    LengthBetween = (MinLengthOf(tempOrigVal, lowerVal) And MaxLengthOf(tempOrigVal, higherVal))
    
    fluentPath = CLng(LengthBetween) & " = " & pPrintSettings.TestName & " Should have length between. "
    fluentPath = fluentPath & "Expected: Greater than " & lowerVal & " less than " & higherVal & "; Actual: " & origVal
    
    Call pPrintSettings.TestCollection.Add(fluentPath)
    'Call pPrintSettings.checkIfPrint(LengthBetween, fluentPath, negateValue)
End Function
