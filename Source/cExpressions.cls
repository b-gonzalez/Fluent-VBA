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

Private pMeta As cMeta
Private pTests As cTests
Private pPrinting As cPrinting
Private iPrint As IPrinting
Private pExprDict As Scripting.Dictionary

Private pTestValue As Variant
Private pFluentPath As String

Private Const NOT_IMPLEMENTED_ERROR As Long = 514

Public Enum flAlgorithm
    flRecursive
    flIterative
End Enum

Public Enum flShould
    Contain_
    EndWith_
    EvaluateTo_
    StartWith_
End Enum

Public Enum flHave
    ErrorDescriptionOf_
    ErrorNumberOf_
    ExactSameElementsAs_
    LengthBetween_
    LengthOf_
    MaxLengthOf_
    MinLengthOf_
    SameElementsAs_
    SameTypeAs_
    SameUniqueElementsAs_
End Enum

Public Enum flBe
    Alphabetic_
    Alphanumeric_
    Between_
    Erroneous_
    EqualTo_
    GreaterThan_
    GreaterThanOrEqualTo_
    IdenticalTo_
    InDataStructure_
    InDataStructures_
    LessThan_
    LessThanOrEqualTo_
    Numeric_
    OneOf_
    Something_
End Enum

Public Property Let TestValue(ByVal value As Variant)
    pTestValue = value
End Property

Public Property Set TestValue(ByVal value As Variant)
    Set pTestValue = value
End Property

Public Property Get TestValue() As Variant
    If IsObject(pTestValue) Then
        Set TestValue = pTestValue
    Else
        TestValue = pTestValue
    End If
End Property

Public Property Set setMeta(ByVal value As cMeta)
    Set pMeta = value
End Property

Public Property Set setTests(ByVal value As cTests)
    Set pTests = value
End Property

Public Property Set setPrinting(ByVal value As IPrinting)
    Set pPrinting = value
    Set iPrint = value
End Property

Public Function InDataStructures( _
ByVal testingValue As Variant, _
ByVal method As flAlgorithm, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant) As Variant
    Dim col As VBA.Collection
    Dim i As Long
    
    InDataStructures = Null
    
    For i = LBound(testingInput) To UBound(testingInput)
        If IsDataStructure(testingInput(i)) Then
            Set col = getNestedElementsRecur(testingInput)
        
            InDataStructures = Me.InDataStructure(testingValue, col, method, negateValue, updateFluentPath:=False)
        End If
    Next i
    
    Call addToTestAndFluentPath("InDataStructures", InDataStructures, CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function InDataStructure( _
ByVal testingValue As Variant, _
ByVal DataStructure As Variant, _
Optional ByVal method As flAlgorithm = flRecursive, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    Dim elem As Variant
    Dim col As Collection
    Dim tempBool As Boolean
    Dim strTestingVal As String
    
    InDataStructure = Null
    
    If IsDataStructure(DataStructure) Then
        tempBool = False
    
        If method = flRecursive Then
            Set col = getNestedElementsRecur(DataStructure)
        ElseIf method = flIterative Then
            Set col = getNestedElementsIter(DataStructure)
        End If
            
        For Each elem In col
            If testingValue = elem Then
                tempBool = True
                Exit For
            End If
        Next elem
        
        InDataStructure = tempBool

    End If
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("InDataStructure", InDataStructure, CBool(negateValue), testingValue, DataStructure)
    End If
    
End Function

Private Function IsDataStructure(ByVal DataStructure As Variant)
    Dim dsTypeName As String
    Dim tempBool As Boolean
    Dim fluentBool As Boolean
    Dim elem As Variant
    Dim tempErr As Long
    
    dsTypeName = TypeName(DataStructure)
    tempBool = False

    On Error Resume Next
        For Each elem In DataStructure
            Exit For
        Next elem
        
        tempErr = Err.Number
        
        tempBool = (tempErr = 0)
        
    On Error GoTo 0
    
    IsDataStructure = tempBool
End Function

Private Function getNestedElementsRecur(ByVal DataStructure As Variant, Optional ByRef col As Collection = Nothing) As Collection
    Dim elem As Variant
    Dim tempDataStructure As Variant
    
    
    If col Is Nothing Then
        Set col = New Collection
    End If
    
    If IsDataStructure(DataStructure) Then
        If TypeName(DataStructure) = "Dictionary" Then
            tempDataStructure = DataStructure.Items
        Else
            If IsObject(DataStructure) Then
                Set tempDataStructure = DataStructure
            Else
                tempDataStructure = DataStructure
            End If
        End If
        
        For Each elem In tempDataStructure
            If IsDataStructure(elem) Then
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

Private Function getNestedElementsIter(ByVal DataStructure As Variant) As Collection
    Dim elem As Variant
    Dim col As Collection
    Dim tempDataStructure As Variant
    Dim i As Long
    Dim hasDataStructures As Boolean

    Set col = New Collection

    If TypeName(DataStructure) = "Dictionary" Then
        tempDataStructure = DataStructure.Items
    Else
        If IsObject(DataStructure) Then
            Set tempDataStructure = DataStructure
        Else
            tempDataStructure = DataStructure
        End If
    End If

    For Each elem In tempDataStructure
        col.Add elem
    Next elem
    
    Do
        hasDataStructures = False
        For i = 1 To col.Count
            If IsDataStructure(col(i)) Then
                hasDataStructures = True
                For Each elem In col(i)
                    col.Add elem
                Next elem
                col.Remove i
                Exit For
            End If
        Next i
    Loop While hasDataStructures = True

    Set getNestedElementsIter = col
End Function

Public Function GreaterThan(ByVal testingValue As Double, ByVal testingInput As Double, Optional ByVal negateValue As Boolean = False) As Variant
    
    GreaterThan = Null
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        GreaterThan = (testingValue > testingInput)
    End If
    
    Call addToTestAndFluentPath("GreaterThan", GreaterThan, CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function GreaterThanOrEqualTo( _
ByVal testingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    
    GreaterThanOrEqualTo = Null
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        GreaterThanOrEqualTo = (testingValue >= testingInput)
    End If
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("GreaterThanOrEqualTo", GreaterThanOrEqualTo, CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function LessThan( _
ByVal testingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    
    LessThan = Null
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        LessThan = (testingValue < testingInput)
    End If
    
    Call addToTestAndFluentPath("LessThan", LessThan, CBool(negateValue), testingValue, testingInput)

End Function

Public Function LessThanOrEqualTo( _
ByVal testingValue As Double, _
ByVal testingInput As Double, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    
    LessThanOrEqualTo = Null
    
    If IsNumeric(testingValue) And IsNumeric(testingInput) Then
        LessThanOrEqualTo = (testingValue <= testingInput)
    End If
       
    If updateFluentPath Then
        Call addToTestAndFluentPath("LessThanOrEqualTo", LessThanOrEqualTo, CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function EqualTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant

    Dim tempBool As Boolean
    
    EqualTo = Null
    
    tempBool = False
    
    If IsNull(testingValue) Or IsNull(testingInput) Then
        tempBool = (TypeName(testingValue) = TypeName(testingInput))
        
        testingValue = TypeName(testingValue)
        testingInput = TypeName(testingInput)
    Else
        tempBool = (testingValue = testingInput)
        
        If pMeta.ApproximateEqual Then
            If TypeName(testingValue) <> TypeName(testingInput) Then
                If IsNumeric(testingValue) And IsNumeric(testingInput) Then
                    tempBool = (Abs(testingValue - testingInput) <= pMeta.Epsilon)
                ElseIf VarType(testingValue) = vbString Or VarType(testingInput) = vbString Then
                    tempBool = (StrComp(CStr(testingValue), CStr(testingInput), vbTextCompare) = 0)
                End If
            End If
        End If
    End If
    
    EqualTo = tempBool
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("EqualTo", EqualTo, CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function Contain( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue As Boolean = False) As Variant

    Contain = Null
        
    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        Contain = CStr(testingValue) Like "*" & CStr(testingInput) & "*"
    End If
        
    Call addToTestAndFluentPath("Contain", Contain, CBool(negateValue), testingValue, testingInput)
        
End Function

Public Function StartWith( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim valLength As Long
    
    StartWith = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        valLength = Len(testingInput)
        StartWith = Left$(testingValue, valLength) = CStr(testingInput)
    End If
    
    Call addToTestAndFluentPath("StartWith", StartWith, CBool(negateValue), testingValue, testingInput)
        
End Function

Public Function EndWith( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim valLength As Long
    
    EndWith = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        valLength = Len(testingInput)
        EndWith = Right$(testingValue, valLength) = CStr(testingInput)
    End If
    
    Call addToTestAndFluentPath("EndWith", EndWith, CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function LengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False) As Variant
    
    LengthOf = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        LengthOf = (Len(CStr(testingValue)) = testingInput)
    End If
    
    Call addToTestAndFluentPath("LengthOf", LengthOf, CBool(negateValue), testingValue, testingInput)
            
End Function

Public Function MaxLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    
    MaxLengthOf = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        MaxLengthOf = (Len(CStr(testingValue)) <= testingInput)
    End If
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("MaxLengthOf", MaxLengthOf, CBool(negateValue), testingValue, testingInput)
    End If
            
End Function

Public Function MinLengthOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue As Boolean = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant
    
    MinLengthOf = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        MinLengthOf = (Len(CStr(testingValue)) >= testingInput)
    End If
        
    If updateFluentPath Then
        Call addToTestAndFluentPath("MinLengthOf", MinLengthOf, CBool(negateValue), testingValue, testingInput)
    End If
        
End Function

Public Function Something(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    
    Something = Null
    
    If IsObject(testingValue) Then
        Something = (Not testingValue Is Nothing)
    End If
    
    Call addToTestAndFluentPath("Something", Something, CBool(negateValue), testingValue)
        
End Function

Public Function Between( _
ByVal testingValue As Variant, _
ByVal LowerVal As Double, _
ByVal HigherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    
    Between = Null
    
    If IsNumeric(testingValue) Then
        Between = (GreaterThanOrEqualTo(testingValue, LowerVal, updateFluentPath:=False) And LessThanOrEqualTo(testingValue, HigherVal, updateFluentPath:=False))
    End If
    
    Call addToTestAndFluentPath("Between", Between, CBool(negateValue), testingValue:=testingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)
End Function

Public Function OneOf(ByVal testingValue As Variant, ByVal negateValue As Boolean, ByVal testingInput As Variant) As Variant
    Dim i As Long
    Dim tempBool As Boolean
    
    OneOf = Null
        
    tempBool = False
    
    For i = LBound(testingInput) To UBound(testingInput)

        If IsObject(testingInput(i)) Or IsObject(testingValue) Then
            If TypeName(testingInput(i)) = TypeName(testingValue) Then
                tempBool = True
            End If
        ElseIf testingInput(i) = testingValue Then
            tempBool = True
        ElseIf IsNull(testingInput(i)) Or IsNull(testingValue) Then
            If TypeName(testingInput(i)) = TypeName(testingValue) Then
                tempBool = True
            End If
        End If
        
        If tempBool Then Exit For
    Next i
    
    OneOf = tempBool
    
    If TypeOf testingValue Is Object  Then testingValue = TypeName(testingValue)
    
    Call addToTestAndFluentPath("OneOf", OneOf, CBool(negateValue), testingValue, testingInput)
        
End Function

Public Function LengthBetween( _
ByVal testingValue As Variant, _
ByVal LowerVal As Double, _
ByVal HigherVal As Double, _
Optional ByVal negateValue As Boolean = False) As Variant
    Dim tempTestVal As Double
    
    LengthBetween = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        tempTestVal = Len(CStr(testingValue))
        LengthBetween = (MinLengthOf(tempTestVal, LowerVal, updateFluentPath:=False) And MaxLengthOf(tempTestVal, HigherVal, updateFluentPath:=False))
    End If
    
    Call addToTestAndFluentPath("LengthBetween", LengthBetween, CBool(negateValue), testingValue:=testingValue, LowerVal:=LowerVal, HigherVal:=HigherVal)

End Function

Public Function EvaluateTo( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue = False) As Variant

    EvaluateTo = Null

    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        EvaluateTo = (Excel.Evaluate(testingValue) = testingInput)
    End If
    
    Call addToTestAndFluentPath("EvaluateTo", EvaluateTo, CBool(negateValue), testingValue, testingInput)
    
End Function

Public Function Alphabetic(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    
    Alphabetic = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        For i = 1 To Len(testingValue)
            tempChar = Mid$(testingValue, i, 1)
            
            If tempChar Like "[A-Za-z]" Then
                tempAlph = tempAlph & tempChar
            End If
        Next i
        
        Alphabetic = (testingValue = tempAlph)
    End If
    
    Call addToTestAndFluentPath("Alphabetic", Alphabetic, CBool(negateValue), testingValue)
End Function

Public Function Numeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempNum As String
    
    Numeric = Null
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        For i = 1 To Len(testingValue)
            tempChar = Mid$(testingValue, i, 1)
            If tempChar Like "[0-9]" Then
                tempNum = tempNum & tempChar
            End If
        Next i
        
        Numeric = (testingValue = tempNum)
    End If
     
    Call addToTestAndFluentPath("Numeric", Numeric, CBool(negateValue), testingValue)
End Function

Public Function Alphanumeric(ByVal testingValue As Variant, Optional ByVal negateValue As Boolean = False) As Variant
    Dim i As Long
    Dim tempChar As String
    Dim tempAlph As String
    Dim tempNum As String
    Dim b As Boolean
    
    Alphanumeric = Null
    
    b = False
    
    If Not IsObject(testingValue) And Not IsArray(testingValue) Then
        For i = 1 To Len(testingValue)
            tempChar = Mid$(testingValue, i, 1)
            If tempChar Like "[A-Za-z]" Then
                tempAlph = tempAlph & tempChar
            ElseIf tempChar Like "[0-9]" Then
                tempNum = tempNum & tempChar
            End If
        Next i
        
        If Len(tempAlph) > 1 And Len(tempNum) > 1 Then
            If Len(tempAlph) + Len(tempNum) = Len(testingValue) Then
                b = True
            End If
        End If
        
        Alphanumeric = b
    End If
    
    Call addToTestAndFluentPath("Alphanumeric", Alphanumeric, CBool(negateValue), testingValue)
End Function

Public Function Erroneous(ByVal testingValue As Variant, Optional ByVal negateValue = False) As Variant
    Dim tempTestingVal As Variant
    Dim tempErrType As Variant
    
    Erroneous = Null
    
    tempTestingVal = False
    
    If TypeOf testingValue Is ErrObject Then
        tempTestingVal = (testingValue.Number <> 0)
    Else
        On Error Resume Next
            tempTestingVal = Excel.Evaluate(testingValue)
            tempTestingVal = (TypeName(tempTestingVal) = "Error")
        On Error GoTo 0
    End If
    
    Erroneous = tempTestingVal
    
    If TypeOf testingValue Is Object  Then
        tempErrType = TypeName(testingValue)
    Else
        tempErrType = testingValue
    End If
    
    Call addToTestAndFluentPath("Erroneous", Erroneous, CBool(negateValue), tempErrType)
End Function

Public Function ErrorNumberOf( _
ByVal testingValue As Variant, _
ByVal testingInput As Long, _
Optional ByVal negateValue = False) As Variant
    Dim temp As Variant
'    Dim e1 As Variant
    
    ErrorNumberOf = Null
    
    On Error Resume Next
        temp = Excel.Evaluate(testingValue)
        Err.Raise (testingInput)
        temp = Err.Number
        ErrorNumberOf = (temp = testingInput)
    On Error GoTo 0
    
    Call addToTestAndFluentPath("ErrorNumberOf", ErrorNumberOf, CBool(negateValue), testingValue, testingInput)
End Function

Public Function ErrorDescriptionOf( _
ByVal testingValue As Variant, _
ByVal testingInput As String, _
Optional ByVal negateValue = False) As Variant
    Dim tempTestVal As Variant
    
    ErrorDescriptionOf = Null
    
    On Error Resume Next
        tempTestVal = CStr(Excel.Evaluate(testingValue))
        tempTestVal = Replace(tempTestVal, "Error ", "")
        Err.Raise tempTestVal
        tempTestVal = Err.Description
        ErrorDescriptionOf = (tempTestVal = testingInput)
    On Error GoTo 0
    
    Call addToTestAndFluentPath("ErrorDescriptionOf", CBool(ErrorDescriptionOf), CBool(negateValue), tempTestVal, testingInput)
End Function

Public Function SameTypeAs( _
ByVal testingValue As Variant, _
ByVal testingInput As Variant, _
Optional ByVal negateValue = False, _
Optional ByVal updateFluentPath As Boolean = False) As Variant

    Dim t1 As String
    Dim t2 As String
    
    SameTypeAs = Null
    
    t1 = TypeName(testingValue)
    t2 = TypeName(testingInput)
    
    SameTypeAs = EqualTo(t1, t2, updateFluentPath:=False)
    
    Call addToTestAndFluentPath("SameTypeAs", SameTypeAs, CBool(negateValue), testingValue, testingInput)
End Function

Public Function IdenticalTo( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant, _
ByVal updateFluentPath As Boolean) As Variant
    
    IdenticalTo = Null
    
    If IsDataStructure(testingInput) And IsDataStructure(testingInput) Then
        IdenticalTo = EqualTo(InputToStringRecur(testingValue), InputToStringRecur(testingInput), updateFluentPath:=False)
    End If
    
    If updateFluentPath Then
        Call addToTestAndFluentPath("IdenticalTo", IdenticalTo, CBool(negateValue), testingValue, testingInput)
    End If

End Function

Public Function ExactSameElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal updateFluentPath As Boolean, _
ByVal testingInput As Variant) As Variant

    Dim elem As Variant
    Dim col As Collection
    Dim col2 As Collection
    
    ExactSameElementsAs = Null
    
    If IsDataStructure(testingValue) And IsDataStructure(testingInput) Then
        Set col = New Collection
        Set col2 = New Collection
        
        For Each elem In testingValue
            col.Add elem
        Next elem
        
        For Each elem In testingInput
            col2.Add elem
        Next elem
        
        ExactSameElementsAs = IdenticalTo(col, negateValue, col2, False)
    End If
    
    Call addToTestAndFluentPath("ExactSameElementsAs", ExactSameElementsAs, CBool(negateValue), testingValue, testingInput)

End Function

Public Function SameUniqueElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant) As Variant

    Dim elem As Variant
    Dim d As Scripting.Dictionary
    Dim tempStr As String
    
    SameUniqueElementsAs = Null
    
    If IsDataStructure(testingValue) And IsDataStructure(testingInput) Then
        Set d = New Scripting.Dictionary
        tempStr = ""
        
        For Each elem In testingValue
            tempStr = ToString(elem)
            If Not d.Exists(tempStr) Then
                d.Add tempStr, Empty
            End If
        Next elem
        
        tempStr = ""
        
        For Each elem In testingInput
            tempStr = ToString(elem)
            If Not d.Exists(tempStr) Then
                d.Add tempStr, Empty
            Else
                d(tempStr) = tempStr
            End If
        Next elem
        
        SameUniqueElementsAs = IdenticalTo(d.Keys, negateValue, d.Items, False)
    End If
    
    Call addToTestAndFluentPath("SameUniqueElementsAs", SameUniqueElementsAs, CBool(negateValue), testingValue, testingInput)

End Function

Public Function SameElementsAs( _
ByVal testingValue As Variant, _
ByVal negateValue As Boolean, _
ByVal testingInput As Variant) As Variant

    Dim i As Long
    Dim elem As Variant
    Dim d As Scripting.Dictionary
    Dim d2 As Scripting.Dictionary
    Dim tempStr As String
    
    SameElementsAs = Null
    
    If IsDataStructure(testingValue) And IsDataStructure(testingInput) Then
        i = 1
        elem = Empty
        Set d = New Scripting.Dictionary
        Set d2 = New Scripting.Dictionary
        tempStr = ""
        
        For Each elem In testingValue
            tempStr = ToString(elem)
            
            If d.Exists(tempStr) Then
                d(tempStr) = d(tempStr) + 1
            Else
                d.Add ToString(elem), i
            End If
        Next elem
        
        tempStr = ""
        
        For Each elem In testingInput
            tempStr = ToString(elem)
            
            If d2.Exists(tempStr) Then
                d2(tempStr) = d2(tempStr) + 1
            Else
                d2.Add ToString(elem), i
            End If
        Next elem
        
        SameElementsAs = IdenticalTo(d, negateValue, d2, False)
    End If
    
    Call addToTestAndFluentPath("SameElementsAs", SameElementsAs, CBool(negateValue), testingValue, testingInput)

End Function

Private Sub addToTestAndFluentPath( _
    ByVal functionName As String, _
    ByVal result As Variant, _
    ByVal negateValue As Boolean, _
    ByVal testingValue As Variant, _
    Optional ByVal testingInput As Variant, _
    Optional LowerVal As Variant, _
    Optional HigherVal As Variant)
    
    Dim test As cTest
    Dim shouldBool As String
    
    Set test = New cTest
    
    If IsNull(result) Then
'        result = TypeName(result)
        test.HasNull = True
    End If

    If negateValue Then
        result = Not result
    End If

    With test
        .functionName = functionName
        .result = result
        .testingInput = Empty
        .LowerVal = Empty
        .HigherVal = Empty
        .StrTestInput = ""
        .StrTestValue = ""
    End With
        
    If TypeName(testingValue) <> "Nothing" Then
        If TypeOf testingValue Is Object  Then
            Set test.testingValue = testingValue
        Else
            test.testingValue = testingValue
        End If
    End If
    
    If Not IsMissing(testingInput) Then
        If TypeName(testingInput) <> "Nothing" Then
            If TypeOf testingInput Is Object  Then
                Set test.testingInput = testingInput
            Else
                test.testingInput = testingInput
            End If
        End If
    End If

    If Not IsMissing(LowerVal) Then
        test.LowerVal = LowerVal
    End If

    If Not IsMissing(HigherVal) Then
        test.HigherVal = HigherVal
    End If
    
    If Not IsMissing(LowerVal) And Not IsMissing(HigherVal) Then
        test.testingInput = Array(CStr(LowerVal), CStr(HigherVal))
    End If
    
    With test
        .StrTestInput = ToString(.testingInput)
        .StrTestValue = ToString(.testingValue)
    End With
    
    
    Call addFluentPathNew(test, negateValue)

End Sub

Function Should_(shouldType As flShould, Optional testingInput As Variant) As String
    Dim shouldStr As String
    
    shouldStr = ""
    
    Select Case shouldType
        Case Contain_
            shouldStr = "contain"
        Case EndWith_
            shouldStr = "endWith"
        Case EvaluateTo_
            shouldStr = "evaluate to"
        Case StartWith_
            shouldStr = "start with"
        Case Else
            Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
    End Select
    
    If Not IsMissing(testingInput) Then
        shouldStr = shouldStr & " " & testingInput
    End If

    Should_ = shouldStr
End Function

Function Be_(beType As flBe, Optional TesingInput As String = "", Optional LowerVal As Variant, Optional HigherVal As Variant)
    Dim beStr As String
    Dim funcStr As String
    
    beStr = "be "
    
    Select Case beType
        Case Alphabetic_
            funcStr = "alphabetic"
        Case Alphanumeric_
            funcStr = "alphanumeric"
        Case Between_
            funcStr = "between"
        Case Erroneous_
            funcStr = "erroneous"
        Case EqualTo_
            funcStr = "equal to"
        Case GreaterThan_
            funcStr = "greater than"
        Case GreaterThanOrEqualTo_
            funcStr = "greater than or equal to"
        Case IdenticalTo_
            funcStr = "identical to"
        Case InDataStructure_
            funcStr = "in data structure"
        Case InDataStructures_
            funcStr = "in data structures"
        Case LessThan_
            funcStr = "less than"
        Case LessThanOrEqualTo_
            funcStr = "less than or equal to"
        Case Numeric_
            funcStr = "numeric"
        Case OneOf_
            funcStr = "one of"
        Case Something_
            funcStr = "something"
        Case Else
            Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
            
    End Select
    
    beStr = beStr & funcStr
    
    If TesingInput <> "" And IsMissing(LowerVal) And IsMissing(HigherVal) Then
        beStr = beStr & " " & TesingInput
    End If
    
    If TesingInput = "" And Not IsMissing(LowerVal) And Not IsMissing(HigherVal) Then
        beStr = beStr & " " & LowerVal & " and " & HigherVal
    End If

    Be_ = beStr
End Function

Function Have_(haveType As flHave, Optional TesingInput As String = "", Optional LowerVal As Variant, Optional HigherVal As Variant)
    Dim haveStr As String
    Dim funcStr As String
    
    haveStr = "have "
    
    Select Case haveType
        Case ErrorDescriptionOf_
            funcStr = "error description of"
        Case ErrorNumberOf_
            funcStr = "error number of"
        Case LengthBetween_
            funcStr = "length between"
        Case LengthOf_
            funcStr = "length of"
        Case MaxLengthOf_
            funcStr = "max length of"
        Case MinLengthOf_
            funcStr = "min length of"
        Case ExactSameElementsAs_
            funcStr = "exact same elements as"
        Case SameElementsAs_
            funcStr = "same elements as"
        Case SameTypeAs_
            funcStr = "same type as"
        Case SameUniqueElementsAs_
            funcStr = "same unique elements as"
        Case Else
            Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
    End Select
    
    haveStr = haveStr & funcStr
    
    If Not IsMissing(TesingInput) And IsMissing(LowerVal) And IsMissing(HigherVal) Then
        haveStr = haveStr & " " & TesingInput
    End If
    
    If IsMissing(TesingInput) And Not IsMissing(LowerVal) And Not IsMissing(HigherVal) Then
        haveStr = haveStr & " " & LowerVal & " and " & HigherVal
    End If

    Have_ = haveStr

End Function

Public Function buildFluentPath(test As cTest, FluentPath As String, tempFluentPath As String) As String
    If IsNull(test.result) Then
        test.result = TypeName(test.result)
    End If

    With test
        If .StrTestInput = "" Then
            tempFluentPath = FluentPath & _
            " Testing value: " & .StrTestValue & ";" & _
            " Expectation: Testing value " & tempFluentPath & ";" & _
            " Result: " & .result
        Else
            tempFluentPath = FluentPath & _
            " Testing value: " & .StrTestValue & ";" & _
            " Testing input: " & .StrTestInput & ";" & _
            " Expectation: Testing value " & tempFluentPath & ";" & _
            " Result: " & .result
        End If
    End With
    
    buildFluentPath = tempFluentPath
End Function

Public Function getFunctionExpr(test As cTest)
    Dim tempFluentPath As String
    
    With test
        Select Case .functionName
            Case "Alphabetic"
                tempFluentPath = Be_(Alphabetic_)
            Case "Alphanumeric"
                tempFluentPath = Be_(Alphanumeric_)
            Case "Between"
                tempFluentPath = Be_(Between_, LowerVal:=.LowerVal, HigherVal:=.HigherVal)
            Case "Contain"
                tempFluentPath = Should_(Contain_, .testingInput)
            Case "Erroneous"
                tempFluentPath = Be_(Erroneous_)
            Case "ErrorNumberOf"
                tempFluentPath = Have_(ErrorNumberOf_, .testingInput)
            Case "ErrorDescriptionOf"
                tempFluentPath = Have_(ErrorDescriptionOf_, .testingInput)
            Case "EqualTo"
                tempFluentPath = Be_(EqualTo_, .testingInput)
            Case "EndWith"
                tempFluentPath = Should_(EndWith_, .testingInput)
            Case "EvaluateTo"
                tempFluentPath = Should_(EvaluateTo_, .StrTestInput)
            Case "ExactSameElementsAs"
                tempFluentPath = Have_(ExactSameElementsAs_, .StrTestInput)
            Case "GreaterThan"
                tempFluentPath = Be_(GreaterThan_, .testingInput)
            Case "GreaterThanOrEqualTo"
                tempFluentPath = Be_(GreaterThanOrEqualTo_, .testingInput)
            Case "IdenticalTo"
                tempFluentPath = Be_(IdenticalTo_, .StrTestInput)
            Case "InDataStructure"
                tempFluentPath = Be_(InDataStructure_, .StrTestInput)
            Case "InDataStructures"
                tempFluentPath = Be_(InDataStructures_, .StrTestInput)
            Case "LengthBetween"
                tempFluentPath = Be_(InDataStructures_, .StrTestInput)
            Case "LengthOf"
                tempFluentPath = Have_(LengthOf_, .testingInput)
            Case "LessThan"
                tempFluentPath = Be_(LessThan_, .testingInput)
            Case "LessThanOrEqualTo"
                tempFluentPath = Be_(LessThanOrEqualTo_, .testingInput)
            Case "MaxLengthOf"
                tempFluentPath = Have_(MaxLengthOf_, .testingInput)
            Case "MinLengthOf"
                tempFluentPath = Have_(MinLengthOf_, .testingInput)
            Case "Numeric"
                tempFluentPath = Be_(Numeric_)
            Case "OneOf"
                tempFluentPath = Be_(OneOf_, .StrTestInput)
            Case "SameElementsAs"
                tempFluentPath = Have_(SameElementsAs_, .StrTestInput)
            Case "SameTypeAs"
                tempFluentPath = Have_(SameTypeAs_, .StrTestInput)
            Case "SameUniqueElementsAs"
                tempFluentPath = Have_(SameUniqueElementsAs_, .StrTestInput)
            Case "Something"
                tempFluentPath = Be_(Something_)
            Case "StartWith"
                tempFluentPath = Should_(StartWith_, .testingInput)
            Case Else
                Err.Raise NOT_IMPLEMENTED_ERROR, "Function not implemented"
        End Select
    End With
    
    getFunctionExpr = tempFluentPath
End Function

Function getFluentPath(FluentPath As String, test As cTest, Optional notShould As Boolean = False)
    Dim LowerVal As Variant
    Dim HigherVal As Variant
    Dim tempFluentPath As String
    
    tempFluentPath = ""
    
    If test.LowerVal <> Empty And test.HigherVal <> Empty Then
        LowerVal = test.LowerVal
        HigherVal = test.HigherVal
    End If

    tempFluentPath = getFunctionExpr(test)
    
    If notShould Then
        tempFluentPath = "should not " & tempFluentPath
    Else
        tempFluentPath = "should " & tempFluentPath
    End If
    
    tempFluentPath = buildFluentPath(test, FluentPath, tempFluentPath)
    
    getFluentPath = tempFluentPath
End Function

Private Sub addFluentPathNew(ByVal test As cTest, ByVal negateValue As Boolean)
    Dim FluentPath As String
    Dim finalFluentPath As String
    Dim fluentElement As Long
    
    FluentPath = ""
    
    If pPrinting.Category <> "" Then
        FluentPath = FluentPath & pPrinting.Category & "="
        fluentElement = 2
    Else
        fluentElement = 1
    End If
    
    FluentPath = FluentPath & getResultNumber(test.result) & "="
    
    If pPrinting.TestName <> "" Then FluentPath = FluentPath & pPrinting.TestName & " "
    
    FluentPath = getFluentPath(FluentPath, test, negateValue)
    
    finalFluentPath = Trim(Split(FluentPath, "=")(fluentElement))
    
    iPrint.FluentPath = finalFluentPath
    
    iPrint.TestCollection.Add FluentPath
    
    test.FluentPath = finalFluentPath
    
    Call pTests.CheckTest(test)
    
End Sub

Private Function ToString(ByVal nv As Variant) As String
    Dim strArgs() As String
    Dim strObj As String
    Dim tempNVTS As String
    Dim i As Long
    Dim elem As Variant
    Dim col As Collection
    Dim counter As Long
    
    i = 0
    counter = 0
    Set col = New Collection
    
    Select Case True
        Case IsArray(nv)
            For Each elem In nv
                    If TypeName(elem) = "Nothing" Then
                        col.Add TypeName(elem)
                    ElseIf TypeOf elem Is Object  Or IsArray(elem) Then
                        If IsDataStructure(elem) Then
                            col.Add InputToStringRecur(elem)
                        ElseIf Not elem Is Nothing Then
                            col.Add TypeName(elem)
                        End If
                    
                    Else
                        col.Add elem
                    End If
                counter = counter + 1
            Next elem
            
            
            If counter > 0 Then
                ReDim strArgs(1 To col.Count)
                    
                For i = 1 To col.Count
                    strArgs(i) = col(i)
                Next i
                
                tempNVTS = Join(strArgs, ", ")
            End If
        Case IsNull(nv)
            tempNVTS = TypeName(nv)
        Case Not TypeOf nv Is Object
            tempNVTS = CStr(nv)
        Case TypeOf nv Is ErrObject
            tempNVTS = TypeName(nv)
        Case TypeOf nv Is Object
            If IsDataStructure(nv) Then
                tempNVTS = InputToStringRecur(nv)
            ElseIf Not nv Is Nothing Then
                tempNVTS = TypeName(nv)
            End If
    End Select
    
    ToString = tempNVTS
        
End Function

Function DatastructureIsEmpty(v As Variant) As Boolean
    Dim elem As Variant
    Dim counter As Long
    Dim b As Boolean
    Dim temp As Variant

    Select Case True
        Case TypeOf v Is Dictionary Or TypeOf v Is Collection
            b = (v.Count = 0)
        
        Case IsArray(v)
            On Error Resume Next
            
                '// will throw an error if v is initialized with no elements
                temp = UBound(v)
                
                b = (Err.Number <> 0)
                
                If Not b Then
                    For Each elem In v
                        If elem <> Empty Then
                            counter = counter + 1
                        End If
                    Next elem
                    
                    b = (counter = 0)
                End If
                
            On Error GoTo 0
    End Select
    
    DatastructureIsEmpty = b
End Function

Private Function InputToStringRecur(v As Variant) As String
    Dim temp As String
    Dim elem As Variant
    Dim emptyDS As Boolean

    temp = ""

    Select Case True
        Case IsArray(v) And DatastructureIsEmpty(v)
            temp = Left(TypeName(v), Len(TypeName(v)) - 2) & "("
        Case IsObject(v) Or IsArray(v)
            If IsArray(v) Then
                temp = Left(TypeName(v), Len(TypeName(v)) - 2) & "("
            Else
                temp = TypeName(v) & "("
            End If
            
            If Not TypeOf v Is Dictionary Then
                For Each elem In v
                    If IsObject(elem) Or IsArray(elem) Then
                        temp = temp & InputToStringRecur(elem) & ", "
                    Else
                        temp = temp & elem & ", "
                    End If
                Next elem
            Else
                For Each elem In v.Keys
                    If IsObject(v(elem)) Or IsArray(v(elem)) Then
                        temp = temp & elem & ": " & InputToStringRecur(v(elem)) & ", "
                    Else
                        temp = temp & elem & ": " & v(elem) & ", "
                    End If
                Next elem
            End If
        Case Else
            temp = temp & elem & ", "
    End Select
    
    If Right(temp, 2) = ", " Then
        temp = Left(temp, Len(temp) - 2)
    End If

    temp = temp & ")"

    InputToStringRecur = temp
End Function

Public Function getResultNumber(result As Variant)
    Dim resultNumber As Variant
    
    resultNumber = Null
    
    If IsNull(result) Then
        resultNumber = 1
    ElseIf TypeName(result) = "Boolean" Then
        resultNumber = CLng(result)
    End If
    
    getResultNumber = resultNumber
End Function

Private Sub Class_Initialize()
    Set pExprDict = New Scripting.Dictionary
End Sub
